# -*- coding: utf-8 -*-
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_模拟.bkp'
TMP = r'D:\<化工工作区>\_probe\panel_all.bkp'

t0 = open(BASE, encoding='utf-8', errors='ignore').read()
SPEC1 = [('3-MP', -1.0), ('NH3', -1.0), ('O2', -1.5), ('3-CP', 1.0), ('H2O', 3.0)]
rows = ['\\ \\ STOIC REACNO = 1 STOIC-CID = "%s" STOIC-SSID = MIXED COEF = %s <0> <0> ' % (c, co)
        for c, co in SPEC1]
newsec = ('? BLOCK RSTOIC "R-101" ? ; "METCBAR_MOLE" ; ; ICON1 ; \n'
          '\\ PARAM TEMP = 405.0 <22> <4> PRES = 1.8 <20> <5> SPEC-OPT = TP \n'
          + '\n'.join(rows) +
          '\\ \\ CONVEX EXT-REACNO = 1 KEY-SSID = MIXED KEY-CID = "3-MP" CONV = .86 <0> <0> \n'
          '\\ \\ PRODUCTS SID = "S-106" \\ \n')

m = re.search(r'\?\s*BLOCK\s+RSTOIC\s+"?R-101\s*"?\s*\?', t0)
nxt = re.search(r'\n\?\s*BLOCK\s', t0[m.start() + 10:])
end = m.start() + 10 + nxt.start() if nxt else len(t0)
open(TMP, 'w', encoding='utf-8', errors='ignore').write(t0[:m.start()] + newsec + t0[end:])
print('patched ->', TMP)

msgs = []


class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            msgs.append(s)


doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
try:
    win32.WithEvents(doc, Sink)
except Exception:
    pass
doc.InitFromArchive2(TMP)
time.sleep(3)
print('InitFromArchive2 done')
print('Engine Ready =', end=' ')
try:
    print(doc.Engine.Ready)
except Exception as ex:
    print('err', ex)

doc.Engine.Run2(False)
t1 = time.time()
while time.time() - t1 < 150:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.2)
    if time.time() - t1 > 8:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(40):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)

print()
print('=== 全部 %d 条控制面板消息 ===' % len(msgs))
for i, x in enumerate(msgs, 1):
    print('%3d| %s' % (i, x[:250]))

print()
print('=== 是否有运行结果 ===')
for b in ['R-101', 'R-501', 'F-501', 'T-201', 'E-601']:
    n = doc.Tree.FindNode(r'\Data\Blocks\%s\Output' % b)
    print('  %-6s %s' % (b, 'None' if n is None else 'ok'))
try:
    doc.Close()
except Exception:
    pass
print('done')
