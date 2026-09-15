# -*- coding: utf-8 -*-
"""最终版：反应与论文第4章一致 + T-201 小再沸器 + PENG-ROB；跑通并保存"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_sim6.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_final.bkp'

t = open(BASE, encoding='utf-8', errors='ignore').read()

# --- 1) R-101 副反应与论文第4章一致 ---
m = re.search(r'\?\s*BLOCK\s+RSTOIC\s+"?R-101\s*"?\s*\?', t)
nxt = re.search(r'\n\?\s*BLOCK\s', t[m.start() + 10:])
end = m.start() + 10 + nxt.start() if nxt else len(t)
NL = '\n'
R101 = NL.join([
    '? BLOCK RSTOIC "R-101" ? ; "METCBAR_MOLE" ; ; ICON1 ; ',
    '\\ PARAM TEMP = 405.0 <22> <4> PRES = 1.8 <20> <5> SPEC-OPT = TP ',
    '\\ \\ STOIC REACNO = 1 STOIC-CID = "3-MP" STOIC-SSID = MIXED COEF = -1.0 <0> <0> ',
    '\\ \\ STOIC REACNO = 1 STOIC-CID = "NH3" STOIC-SSID = MIXED COEF = -1.0 <0> <0> ',
    '\\ \\ STOIC REACNO = 1 STOIC-CID = "O2" STOIC-SSID = MIXED COEF = -1.5 <0> <0> ',
    '\\ \\ STOIC REACNO = 2 STOIC-CID = "3-MP" STOIC-SSID = MIXED COEF = -1.0 <0> <0> ',
    '\\ \\ STOIC REACNO = 2 STOIC-CID = "O2" STOIC-SSID = MIXED COEF = -7.75 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 1 STOIC-CID1 = "3-CP" STOIC-SSID1 = MIXED COEF1 = 1.0 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 1 STOIC-CID1 = "H2O" STOIC-SSID1 = MIXED COEF1 = 3.0 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 2 STOIC-CID1 = "CO2" STOIC-SSID1 = MIXED COEF1 = 6.0 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 2 STOIC-CID1 = "N2" STOIC-SSID1 = MIXED COEF1 = 0.5 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 2 STOIC-CID1 = "H2O" STOIC-SSID1 = MIXED COEF1 = 3.5 <0> <0> ',
    '\\ \\ CONVEX EXT-REACNO = 1 KEY-SSID = MIXED KEY-CID = "3-MP" CONV = .86 <0> <0> ',
    '\\ \\ CONVEX EXT-REACNO = 2 KEY-SSID = MIXED KEY-CID = "3-MP" CONV = .02 <0> <0> ',
    '\\ \\ PRODUCTS SID = "S-106" \\ ',
])
t = t[:m.start()] + R101 + NL + t[end:]

# --- 2) T-201 小再沸器 ---
t = t.replace('"COL-CONFIG" CONDENSER = NONE REBOILER = NONE',
              '"COL-CONFIG" CONDENSER = NONE REBOILER = KETTLE')
mm = re.search(r'"COL-SPECS" BASIS-RDV = 1\.0 <0> <0> ', t)
t = t[:mm.start()] + '"COL-SPECS" BASIS-RDV = 1.0 <0> <0> BASIS-BR = 0.1 <-1> <0> ' + t[mm.end():]

# --- 3) 物性方法 ---
m3 = re.search(r'PARAM\s+BASE\s*=\s*[^\s\\]+', t)
t = t[:m3.start()] + 'PARAM BASE = "PENG-ROB"' + t[m3.end():]

open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)
print('已生成:', OUT, flush=True)

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

print()
print('=== 控制面板尾部 ===', flush=True)
for x in MSGS[-20:]:
    print('   |', x[:190], flush=True)


def g(p):
    n = doc.Tree.FindNode(p)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return 'ERR'


print()
print('=== 模块结果 ===', flush=True)
for b in ['T-201', 'T-301', 'T-401', 'T-402', 'T-403', 'T-404', 'E-601']:
    bb = r'\Data\Blocks\%s\Output' % b
    d = {f: g(bb + '\\' + f) for f in
         ['COND_DUTY', 'REB_DUTY', 'MOLE_RR', 'BOTTOM_TEMP', 'TOP_TEMP', 'B-PRES']}
    print('  %-7s %s' % (b, ' '.join('%s=%s' % (k, v) for k, v in d.items() if v not in (None, ''))),
          flush=True)

print()
print('=== 保存 ===', flush=True)
for p in [r'D:\<化工工作区>\NA-Chemical-10000t_final.bkp',
          r'D:\<化工工作区>\NA-Chemical-10000t_final.apwz']:
    try:
        doc.SaveAs(p)
        print('  已保存:', p, os.path.getsize(p), flush=True)
    except Exception as ex:
        print('  失败 %s: %s' % (p, ex), flush=True)

try:
    doc.Close()
except Exception:
    pass
print('DONE', flush=True)
