# -*- coding: utf-8 -*-
"""S5：修正 T-302 产品板 -> 塔底；跑通并实测三效蒸发链相态"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_s3.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_s5.bkp'
t = open(SRC, encoding='utf-8', errors='ignore').read()
print('源大小', len(t))


def repl_block(x, bid, fn):
    starts = [s.start() for s in re.finditer(r'\?\s*BLOCK\s+', x)]
    for k in range(len(starts)):
        s0 = starts[k]
        s1 = starts[k + 1] if k + 1 < len(starts) else len(x)
        m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', x[s0:s1])
        if m and m.group(2) == bid:
            return x[:s0] + fn(x[s0:s1]) + x[s1:]
    raise RuntimeError(bid)


T302_NEW = ('? BLOCK RADFRAC "T-302" ? ; "METCBAR_MOLE" ; ; ABSBR1 ; \\ PARAM NSTAGE = 8\n'
            'NSTAGEMAX = 9 \\ \\ PARAM2 \\ \\ "COL-CONFIG" CONDENSER = NONE REBOILER =\n'
            'KETTLE \\ \\ FEEDS FEED-SID = "S-113" FEED-STAGE = 1 \\ \\ PRODUCTS\n'
            'PROD-STREAM = "S-212" PROD-STAGE = 1 PROD-PHASE = V P-S = N /  PROD-STREAM\n'
            '= "S-213" PROD-STAGE = 8 PROD-PHASE = L P-S = N \\ \\ "P-SPEC2" PRES1 = 1.0\n'
            '<20> <5> \\ \\ "COL-SPECS" BASIS-RDV = 1.0 <0> <0> BASIS-BR = 0.2 <-1> <0> \\\n'
            '\\ T-EST TEMP-STAGE = 1 TEMP-EST = 40.0 <22> <4> /  TEMP-STAGE = 8 TEMP-EST\n'
            '= 100.0 <22> <4> \\ \\ "KLL-VECS" \\ \\ "TRSZ-VECS" \\ \\ "PCKSR-VECS" \\ \n')

t = repl_block(t, 'T-302', lambda s: T302_NEW)
i = t.find('BLOCK RADFRAC "T-302"')
print('1) T-302 重建 ✓  PROD-STAGE=8, BASIS-BR=0.2')

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
win32.WithEvents(doc, Sink)
doc.InitFromArchive2(OUT)
time.sleep(3)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 480:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 10:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(80):
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


L = []
L.append('=== 面板消息 (%d) ===' % len(MSGS))
for i2, x in enumerate(MSGS):
    L.append('%3d | %s' % (i2, x[:170]))

L.append('')
L.append('=== 流股结果（相态 / 温度 / 质量流量）===')
L.append('  %-8s %-10s %-8s %-10s' % ('流股', 'VFRAC', 'TEMP_C', 'MASS_kgh'))
for s in ['S-101', 'S-102', 'S-103', 'S-126', 'S-127', 'S-128', 'S-104', 'S-105', 'S-106',
          'S-129', 'S-108', 'S-107', 'S-109', 'S-110', 'S-112', 'S-113', 'S-114', 'S-114A',
          'S-115', 'S-116', 'S-116A', 'S-117', 'S-118', 'S-119', 'S-120', 'S-121', 'S-122',
          'S-201', 'S-210', 'S-202', 'S-203', 'S-211', 'S-212', 'S-213',
          'S-401', 'S-402', 'S-214', 'S-205', 'S-403', 'S-206', 'S-404', 'S-207',
          'S-405', 'S-208', 'S-406', 'S-209']:
    L.append('  %-8s %-10s %-8s %-10s' % (
        s, g(r'\Data\Streams\%s\Output\VFRAC\MIXED' % s),
        g(r'\Data\Streams\%s\Output\TEMP\MIXED' % s),
        g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s)))

L.append('')
L.append('=== 后处理段关键组分（S-209 干品 / S-405 母液 / S-213 废水）===')
for s in ['S-209', 'S-405', 'S-213', 'S-203']:
    nn = doc.Tree.FindNode(r'\Data\Streams\%s\Output\MASSFLOW3' % s)
    if nn is None:
        L.append('  %s: 无数据' % s)
        continue
    L.append('  --- %s ---' % s)
    try:
        for i3 in range(nn.Elements.Count):
            e = nn.Elements.Item(i3)
            try:
                if e.Value and abs(float(e.Value)) > 0.01:
                    L.append('      %-8s %9.2f' % (e.Name, float(e.Value)))
            except Exception:
                pass
    except Exception as ex:
        L.append('      读取失败 %s' % ex)

try:
    doc.SaveAs(OUT.replace('.bkp', '_s.bkp'))
    L.append('')
    L.append('已保存')
except Exception as ex:
    L.append('保存失败: %s' % ex)
try:
    doc.Close()
except Exception:
    pass

open(r'D:\<化工工作区>\_probe\s5.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('DONE')
