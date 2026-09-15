# -*- coding: utf-8 -*-
"""P1：修 T-401 采出比 + 把 S-111 补充甲苯降到合理值"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_v19.bkp'

t = open(SRC, encoding='utf-8', errors='ignore').read()

# ---- 1) 打印 S-111 段原文 ----
m = re.search(r'\?\s*STREAM\s+MATERIAL\s+"S-111"\s*\?', t)
print('S-111 段头位置:', m.start() if m else None)
if m:
    i = m.start()
    j = t.find('? STREAM', i + 20)
    if j < 0:
        j = t.find('? \\nSTREAM', i + 20)
    print('段原文:')
    print(repr(t[i:i + 420]))
print()

MAKEUP = 0.25   # kmol/h

# ---- 2) 改 S-111 流量 ----
seg_start = m.start()
nxt = re.search(r'\?\s*STREAM\s+MATERIAL\s+"S-1', t[seg_start + 30:])
seg_end = seg_start + 30 + nxt.start() if nxt else seg_start + 700
seg = t[seg_start:seg_end]
seg2 = re.sub(r'(TOTAL\s*=\s*)24\.47', r'\g<1>%s' % MAKEUP, seg)
seg2 = re.sub(r'(CID\s*=\s*TOL\s+FLOW\s*=\s*)24\.47', r'\g<1>%s' % MAKEUP, seg2)
print('S-111 TOTAL 替换:', seg2 != seg)
t = t[:seg_start] + seg2 + t[seg_end:]

# ---- 3) T-401：D:F=0.83，回流比提高 ----
starts = [s.start() for s in re.finditer(r'\?\s*BLOCK\s+', t)]
for k in range(len(starts) - 1, -1, -1):
    s0 = starts[k]
    s1 = starts[k + 1] if k + 1 < len(starts) else len(t)
    mm = re.match(r'\?\s*BLOCK\s+RADFRAC\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
    if mm and mm.group(1) == 'T-401':
        sg = t[s0:s1]
        sg2 = re.sub(r'D:F = [\d.]+ <-1> <0>', 'D:F = 0.8300 <-1> <0>', sg, count=1)
        sg2 = re.sub(r'BASIS-RR = [\d.]+ <-1> <0>', 'BASIS-RR = 3.0 <-1> <0>', sg2, count=1)
        print('T-401 段改动:', sg2 != sg)
        t = t[:s0] + sg2 + t[s1:]
        break

# ---- 4) 物性方法保持 NRTL（v19 已设）----
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_p1.bkp'
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)
print('写出:', OUT, os.path.getsize(OUT))

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
doc.InitFromArchive2(OUT)
time.sleep(3)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 420:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 10:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(60):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)


def g(p):
    n = doc.Tree.FindNode(p)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return None


print()
print('=== 面板尾部 ===')
for x in MSGS[-16:]:
    print('   |', x[:170])
print()
print('=== 塔 ===')
for b in ['T-401', 'T-402', 'T-403', 'T-404']:
    bb = r'\Data\Blocks\%s\Output' % b
    print('  %-6s COND=%-11s REB=%-11s RR=%-7s Ttop=%-9s Tbot=%-9s' % (
        b, g(bb + r'\COND_DUTY'), g(bb + r'\REB_DUTY'), g(bb + r'\MOLE_RR'),
        g(bb + r'\TOP_TEMP'), g(bb + r'\BOTTOM_TEMP')))
print()
print('=== 甲苯相关流股 ===')
for s in ['S-111', 'S-114', 'S-115', 'S-116', 'S-121']:
    b = r'\Data\Streams\%s\Output' % s
    print('  --- %s MASS=%s' % (s, g(b + r'\MASSFLMX\MIXED')))
    nn = doc.Tree.FindNode(b + r'\MASSFLOW3')
    if nn is not None:
        try:
            for i2 in range(nn.Elements.Count):
                e = nn.Elements.Item(i2)
                try:
                    if e.Value and abs(float(e.Value)) > 0.05:
                        print('        %-8s %9.2f' % (e.Name, float(e.Value)))
                except Exception:
                    pass
        except Exception:
            pass
try:
    doc.SaveAs(OUT.replace('.bkp', '_s.bkp'))
    print('已保存')
except Exception as ex:
    print('保存失败:', ex)
try:
    doc.Close()
except Exception:
    pass
print('DONE')
