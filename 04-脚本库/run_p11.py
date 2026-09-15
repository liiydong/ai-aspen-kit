# -*- coding: utf-8 -*-
"""P11：生成最终模型（T-402 D:F=0.040 RR=15）并输出完整结果表"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\_probe\p10b.bkp'
FB = r'D:\<化工工作区>\NA-Chemical-10000t_最终版.bkp'
FA = r'D:\<化工工作区>\NA-Chemical-10000t_最终版.apwz'

t = open(SRC, encoding='utf-8', errors='ignore').read()
print('重复性检查: RADFRAC=%d T-401=%d S-111=%d' % (
    len(re.findall(r'\?\s*BLOCK\s+RADFRAC', t)),
    len(re.findall(r'BLOCK\s+RADFRAC\s*\n?\s*"T-401"', t)),
    len(re.findall(r'\?\s*STREAM\s+MATERIAL\s*\n?\s*"S-111"', t))))

MSGS = []


class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            MSGS.append(s)


doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
try:
    win32.WithEvents(doc, Sink)
except Exception:
    pass
doc.InitFromArchive2(SRC)
time.sleep(3)


def g(p):
    n = doc.Tree.FindNode(p)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return None


print()
print('=== 参数回读 ===')
for b, f in [('T-401', 'D:F'), ('T-401', 'BASIS_RR'), ('T-401', 'PRES1'), ('T-402', 'PRES1'),
             ('T-403', 'PRES1'), ('T-404', 'PRES1')]:
    print('  %s %s = %s' % (b, f, g(r'\Data\Blocks\%s\Input\%s' % (b, f))))

doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 320:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 10:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(50):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)

print()
print('=== 错误汇总 ===')
for x in MSGS:
    if 'Errors' in x or 'Warnings' in x or 'completed' in x:
        print('   |', x[:170])
print()
print('=== 塔结果 ===')
hdr = '塔'.ljust(8)
print('  塔      冷凝kW     再沸kW     回流比   塔顶℃     塔底℃')
for b in ['T-201', 'T-401', 'T-402', 'T-403', 'T-404']:
    bb = r'\Data\Blocks\%s\Output' % b
    print('  %-7s %-11s %-11s %-8s %-9s %-9s' % (
        b, g(bb + r'\COND_DUTY'), g(bb + r'\REB_DUTY'), g(bb + r'\MOLE_RR'),
        g(bb + r'\TOP_TEMP'), g(bb + r'\BOTTOM_TEMP')))
print()
print('=== 流股结果 ===')
for s in ['S-104', 'S-106', 'S-109', 'S-110', 'S-113', 'S-114', 'S-115', 'S-116',
          'S-117', 'S-118', 'S-119', 'S-120', 'S-121', 'S-122']:
    b = r'\Data\Streams\%s\Output' % s
    print('  %-7s MASS=%-11s T=%s' % (s, g(b + r'\MASSFLMX\MIXED'), g(b + r'\TEMP_OUT')))
    nn = doc.Tree.FindNode(b + r'\MASSFLOW3')
    if nn is not None:
        for i2 in range(nn.Elements.Count):
            e = nn.Elements.Item(i2)
            try:
                if e.Value and abs(float(e.Value)) > 0.05:
                    print('          %-8s %9.2f' % (e.Name, float(e.Value)))
            except Exception:
                pass

for p in [FA, FB, FB.replace('.bkp', '_s.bkp')]:
    try:
        doc.SaveAs(p)
        print('已保存:', p, os.path.getsize(p))
    except Exception as ex:
        print('保存失败:', ex)
try:
    doc.Close()
except Exception:
    pass
print('DONE')
