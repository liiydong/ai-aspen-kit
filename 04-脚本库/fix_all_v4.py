# -*- coding: utf-8 -*-
"""v4：反应物=STOIC / 产物=STOIC1；四塔补第二条规格 D:F"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_sim3.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_sim4.bkp'
NL = '\n'


def sec_rstoic(bid, temp, pres, react, prod, convex, out):
    """react: [(rn,cid,mag)]; prod: [(rn,cid,mag)]; 系数写正值+显式符号"""
    L = ['? BLOCK RSTOIC "%s" ? ; "METCBAR_MOLE" ; ; ICON1 ; ' % bid,
         '\\ PARAM TEMP = %s <22> <4> PRES = %s <20> <5> SPEC-OPT = TP ' % (temp, pres)]
    for rn, cid, m in react:
        L.append('\\ \\ STOIC REACNO = %d STOIC-CID = "%s" STOIC-SSID = MIXED '
                 'COEF = -%s <0> <0> ' % (rn, cid, m))
    for rn, cid, m in prod:
        L.append('\\ \\ STOIC1 REACNO1 = %d STOIC-CID1 = "%s" STOIC-SSID1 = MIXED '
                 'COEF1 = %s <0> <0> ' % (rn, cid, m))
    for rn, key, cv in convex:
        L.append('\\ \\ CONVEX EXT-REACNO = %d KEY-SSID = MIXED KEY-CID = "%s" '
                 'CONV = %s <0> <0> ' % (rn, key, cv))
    L.append('\\ \\ PRODUCTS SID = "%s" \\ ' % out)
    return NL.join(L)


R101 = sec_rstoic('R-101', '405.0', '1.8',
                  [(1, '3-MP', '1.0'), (1, 'NH3', '1.0'), (1, 'O2', '1.5'),
                   (2, '3-MP', '1.0'), (2, 'O2', '7.5')],
                  [(1, '3-CP', '1.0'), (1, 'H2O', '3.0'),
                   (2, 'CO2', '6.0'), (2, 'H2O', '3.0'), (2, 'HCN', '1.0')],
                  [(1, '3-MP', '.86'), (2, '3-MP', '.02')], 'S-106')

R501 = sec_rstoic('R-501', '140.0', '4.0',
                  [(1, '3-CP', '1.0'), (1, 'H2O', '1.0'),
                   (2, 'NAM', '1.0'), (2, 'H2O', '1.0')],
                  [(1, 'NAM', '1.0'),
                   (2, 'NAC', '1.0'), (2, 'NH3', '1.0')],
                  [(1, '3-CP', '.985'), (2, 'NAM', '.005')], 'S-201')

F501 = sec_rstoic('F-501', '60.0', '1.0',
                  [(1, 'NAOH', '2.0'), (1, 'H2SO4', '1.0')],
                  [(1, 'NA2SO4', '1.0'), (1, 'H2O', '2.0')],
                  [(1, 'H2SO4', '1.0')], 'S-301')

NEW = {'R-101': R101, 'R-501': R501, 'F-501': F501}

# 四塔补第二条规格
DF = {'T-401': '0.50', 'T-402': '0.50', 'T-403': '0.10', 'T-404': '0.30'}
t = open(BASE, encoding='utf-8', errors='ignore').read()

starts = [m.start() for m in re.finditer(r'\?\s*BLOCK\s+[A-Z0-9]+\s+"?[A-Za-z0-9-]+"?\s*\?', t)]
rep = []
for idx in range(len(starts) - 1, -1, -1):
    s0 = starts[idx]
    s1 = starts[idx + 1] if idx + 1 < len(starts) else len(t)
    m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"([^"]+)"\s*\?', t[s0:s1])
    if not m:
        continue
    bid = m.group(2)
    if bid in NEW:
        t = t[:s0] + NEW[bid] + NL + t[s1:]
        rep.append(bid)
    elif bid in DF:
        seg = t[s0:s1]
        if 'BASIS-RR' in seg and 'D:F' not in seg:
            seg2 = seg.replace('"COL-SPECS" BASIS-RDV',
                               '"COL-SPECS" D:F = %s <-1> <0> BASIS-RDV' % DF[bid], 1)
            t = t[:s0] + seg2 + t[s1:]
            rep.append(bid + '(D:F)')
print('已改:', rep, flush=True)
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)

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
print('Ready =', doc.Engine.Ready, flush=True)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 300:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 8:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(40):
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
    doc.SaveAs(r'D:\<化工工作区>\_probe\sim4_saved.bkp')
except Exception:
    pass
try:
    doc.Close()
except Exception:
    pass
print('DONE', flush=True)
