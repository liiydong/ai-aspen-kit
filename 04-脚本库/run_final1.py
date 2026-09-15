# -*- coding: utf-8 -*-
"""最终修复：T-201 加小再沸器 + 物性方法改 NRTL-RK，跑通全流程并读结果"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_sim6.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_conv.bkp'

t = open(BASE, encoding='utf-8', errors='ignore').read()

n1 = t.count('"COL-CONFIG" CONDENSER = NONE REBOILER = NONE')
print('COL-CONFIG(NONE/NONE) 出现次数:', n1)
t = t.replace('"COL-CONFIG" CONDENSER = NONE REBOILER = NONE',
              '"COL-CONFIG" CONDENSER = NONE REBOILER = KETTLE')
m = re.search(r'"COL-SPECS" BASIS-RDV = 1\.0 <0> <0> ', t)
print('COL-SPECS 匹配:', bool(m))
if m:
    t = (t[:m.start()] + '"COL-SPECS" BASIS-RDV = 1.0 <0> <0> BASIS-BR = 0.1 <-1> <0> '
         + t[m.end():])

m2 = re.search(r'PARAM\s+BASE\s*=\s*NRTL\b', t)
print('BASE 匹配:', bool(m2), repr(t[m2.start():m2.end()]) if m2 else '')
if m2:
    t = t[:m2.start()] + 'PARAM BASE = "NRTL-RK"' + t[m2.end():]

open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)
print('写出:', OUT, flush=True)

MSGS = []


class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            MSGS.append(s)


doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
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
for _ in range(80):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)

print()
print('=== 控制面板 %d 条 ===' % len(MSGS), flush=True)
for i, x in enumerate(MSGS, 1):
    print('%3d| %s' % (i, x[:200]), flush=True)


def g(p):
    n = doc.Tree.FindNode(p)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return 'ERR'


print()
print('=== 关键流股结果 ===', flush=True)
print('%-8s %10s %10s %14s %12s' % ('stream', 'T(C)', 'P(bar)', 'MOLE kmol/h', 'MASS kg/h'))
for s in ['S-104', 'S-106', 'S-108', 'S-109', 'S-110', 'S-112', 'S-113', 'S-114',
          'S-115', 'S-116', 'S-117', 'S-118', 'S-119', 'S-120', 'S-121', 'S-122',
          'S-201', 'S-301', 'S-401', 'S-402']:
    b = r'\Data\Streams\%s\Output' % s
    print('%-8s %10s %10s %14s %12s' % (s, g(b + r'\TEMP'), g(b + r'\PRES'),
                                         g(b + r'\MOLE-FLOW'), g(b + r'\MASS-FLOW')))

print()
print('=== 模块热负荷 ===', flush=True)
for b in ['E-101', 'E-104', 'E-201', 'T-201', 'T-401', 'T-402', 'T-403', 'T-404',
          'R-101', 'R-501', 'F-501', 'E-601']:
    bb = r'\Data\Blocks\%s\Output' % b
    print('  %-7s T=%s P=%s DUTY=%s COND=%s REB=%s' % (
        b, g(bb + r'\TEMP'), g(bb + r'\PRES'), g(bb + r'\DUTY'),
        g(bb + r'\COND_DUTY'), g(bb + r'\REB_DUTY')))

try:
    doc.SaveAs(r'D:\<化工工作区>\_probe\final_saved.bkp')
except Exception:
    pass
try:
    doc.Close()
except Exception:
    pass
print('DONE', flush=True)
