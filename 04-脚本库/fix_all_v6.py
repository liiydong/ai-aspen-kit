# -*- coding: utf-8 -*-
"""v6：修塔压单位行号 + T-201 补 FEED-CONVE2 + 物性方法改 NRTL-RK（避开 NaOH 缺 DHVLWT）"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_sim5.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_sim6.bkp'

t = open(BASE, encoding='utf-8', errors='ignore').read()
print('<20> <10> 出现次数:', t.count('<20> <10>'))
print('BASE = NRTL 出现次数:', t.count('PARAM BASE = NRTL'))

t = t.replace('<20> <10>', '<20> <5>')
t = t.replace('PARAM BASE = NRTL', 'PARAM BASE = "NRTL-RK"')
t = t.replace(
    'FEEDS FEED-SID = "S-107" FEED-STAGE = 1 / FEED-SID = "S-108" FEED-STAGE = 6',
    'FEEDS FEED-SID = "S-107" FEED-STAGE = 1 FEED-CONVE2 = "ON-STAGE" / '
    'FEED-SID = "S-108" FEED-STAGE = 6 FEED-CONVE2 = "ON-STAGE"')
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
while time.time() - t0 < 360:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 8:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(60):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)

print()
print('=== 控制面板 %d 条 ===' % len(MSGS), flush=True)
for i, x in enumerate(MSGS, 1):
    print('%3d| %s' % (i, x[:220]), flush=True)

print()
print('=== 结果抽查 ===', flush=True)
for b in ['M-101', 'E-101', 'R-101', 'T-201', 'T-301', 'T-401', 'T-402', 'T-403',
          'T-404', 'R-501', 'F-501', 'E-601']:
    n = doc.Tree.FindNode(r'\Data\Blocks\%s\Output' % b)
    if n is None:
        print('  %-6s 无输出' % b, flush=True); continue
    vals = []
    for f in ['TEMP', 'PRES', 'DUTY', 'B-PRES', 'MOLE_RR', 'COND_DUTY', 'REB_DUTY']:
        try:
            x = n.Elements.Item(f)
            if x is not None and x.Value not in (None, ''):
                vals.append('%s=%s' % (f, x.Value))
        except Exception:
            pass
    print('  %-6s %s' % (b, ', '.join(vals) if vals else '(空)'), flush=True)

try:
    doc.SaveAs(r'D:\<化工工作区>\_probe\sim6_saved.bkp')
except Exception:
    pass
try:
    doc.Close()
except Exception:
    pass
print('DONE', flush=True)
