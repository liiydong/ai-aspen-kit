# -*- coding: utf-8 -*-
"""v5：修正 R-101 副反应配平（5 CO2 + 6.5 O2），其余不动；跑一次看是否收敛"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_sim4.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_sim5.bkp'
NL = '\n'

R101_OLD = r'\?\s*BLOCK\s+RSTOIC\s+"?R-101\s*"?\s*\?'

t = open(BASE, encoding='utf-8', errors='ignore').read()
m = re.search(R101_OLD, t)
nxt = re.search(r'\n\?\s*BLOCK\s', t[m.start() + 10:])
end = m.start() + 10 + nxt.start() if nxt else len(t)

NEW = NL.join([
    '? BLOCK RSTOIC "R-101" ? ; "METCBAR_MOLE" ; ; ICON1 ; ',
    '\\ PARAM TEMP = 405.0 <22> <4> PRES = 1.8 <20> <5> SPEC-OPT = TP ',
    '\\ \\ STOIC REACNO = 1 STOIC-CID = "3-MP" STOIC-SSID = MIXED COEF = -1.0 <0> <0> ',
    '\\ \\ STOIC REACNO = 1 STOIC-CID = "NH3" STOIC-SSID = MIXED COEF = -1.0 <0> <0> ',
    '\\ \\ STOIC REACNO = 1 STOIC-CID = "O2" STOIC-SSID = MIXED COEF = -1.5 <0> <0> ',
    '\\ \\ STOIC REACNO = 2 STOIC-CID = "3-MP" STOIC-SSID = MIXED COEF = -1.0 <0> <0> ',
    '\\ \\ STOIC REACNO = 2 STOIC-CID = "O2" STOIC-SSID = MIXED COEF = -6.5 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 1 STOIC-CID1 = "3-CP" STOIC-SSID1 = MIXED COEF1 = 1.0 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 1 STOIC-CID1 = "H2O" STOIC-SSID1 = MIXED COEF1 = 3.0 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 2 STOIC-CID1 = "CO2" STOIC-SSID1 = MIXED COEF1 = 5.0 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 2 STOIC-CID1 = "H2O" STOIC-SSID1 = MIXED COEF1 = 3.0 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 2 STOIC-CID1 = "HCN" STOIC-SSID1 = MIXED COEF1 = 1.0 <0> <0> ',
    '\\ \\ CONVEX EXT-REACNO = 1 KEY-SSID = MIXED KEY-CID = "3-MP" CONV = .86 <0> <0> ',
    '\\ \\ CONVEX EXT-REACNO = 2 KEY-SSID = MIXED KEY-CID = "3-MP" CONV = .02 <0> <0> ',
    '\\ \\ PRODUCTS SID = "S-106" \\ ',
])
t = t[:m.start()] + NEW + NL + t[end:]
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
    doc.SaveAs(r'D:\<化工工作区>\_probe\sim5_saved.bkp')
except Exception:
    pass
try:
    doc.Close()
except Exception:
    pass
print('DONE', flush=True)
