# -*- coding: utf-8 -*-
"""按原生格式重修：3 个反应器（plain STOIC）+ T-201（吸收塔原生写法）+ T-301（3 进料不同板）"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_模拟.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_sim2.bkp'
NL = '\n'


def sec_rstoic(bid, temp, pres, stoic, convex, prod):
    L = ['? BLOCK RSTOIC "%s" ? ; "METCBAR_MOLE" ; ; ICON1 ; ' % bid,
         '\\ PARAM TEMP = %s <22> <4> PRES = %s <20> <5> SPEC-OPT = TP ' % (temp, pres)]
    for rn, cid, co in stoic:
        L.append('\\ \\ STOIC REACNO = %d STOIC-CID = "%s" STOIC-SSID = MIXED '
                 'COEF = %s <0> <0> ' % (rn, cid, co))
    for rn, key, cv in convex:
        L.append('\\ \\ CONVEX EXT-REACNO = %d KEY-SSID = MIXED KEY-CID = "%s" '
                 'CONV = %s <0> <0> ' % (rn, key, cv))
    L.append('\\ \\ PRODUCTS SID = "%s" \\ ' % prod)
    return NL.join(L)


R101 = sec_rstoic('R-101', '405.0', '1.8',
                  [(1, '3-MP', -1.0), (1, 'NH3', -1.0), (1, 'O2', -1.5),
                   (1, '3-CP', 1.0), (1, 'H2O', 3.0),
                   (2, '3-MP', -1.0), (2, 'O2', -7.5), (2, 'CO2', 6.0),
                   (2, 'H2O', 3.0), (2, 'HCN', 1.0)],
                  [(1, '3-MP', '.86'), (2, '3-MP', '.02')], 'S-106')

R501 = sec_rstoic('R-501', '140.0', '4.0',
                  [(1, '3-CP', -1.0), (1, 'H2O', -1.0), (1, 'NAM', 1.0),
                   (2, 'NAM', -1.0), (2, 'H2O', -1.0), (2, 'NAC', 1.0), (2, 'NH3', 1.0)],
                  [(1, '3-CP', '.985'), (2, 'NAM', '.005')], 'S-201')

F501 = sec_rstoic('F-501', '60.0', '1.0',
                  [(1, 'NAOH', -2.0), (1, 'H2SO4', -1.0), (1, 'NA2SO4', 1.0), (1, 'H2O', 2.0)],
                  [(1, 'H2SO4', '1.0')], 'S-301')

# T-201 吸收塔：照原生 absorber 格式（只有 BASIS-RDV 一条规格）
T201 = NL.join([
    '? BLOCK RADFRAC "T-201" ? ; "METCBAR_MOLE" ; ; ABSBR1 ; ',
    '\\ PARAM NSTAGE = 6 NSTAGEMAX = 7 ',
    '\\ \\ PARAM2 ',
    '\\ \\ "COL-CONFIG" CONDENSER = NONE REBOILER = NONE ',
    '\\ \\ FEEDS FEED-SID = "S-107" FEED-STAGE = 1 /  FEED-SID = "S-108" FEED-STAGE = 6 ',
    '\\ \\ PRODUCTS PROD-STREAM = "S-109" PROD-STAGE = 1 PROD-PHASE = V P-S = N / '
    ' PROD-STREAM = "S-110" PROD-STAGE = 6 PROD-PHASE = L P-S = N ',
    '\\ \\ "P-SPEC2" PRES1 = 1.0 <20> <10> ',
    '\\ \\ "COL-SPECS" BASIS-RDV = 1.0 <0> <0> ',
    '\\ \\ T-EST TEMP-STAGE = 1 TEMP-EST = 40.0 <22> <4> / '
    'TEMP-STAGE = 6 TEMP-EST = 60.0 <22> <4> ',
    '\\ \\ "KLL-VECS" ',
    '\\ \\ "TRSZ-VECS" ',
    '\\ \\ "PCKSR-VECS" \\ ',
])

# T-301 萃取塔：3 条进料分板（每板只允许一条）
T301 = NL.join([
    '? BLOCK EXTRACT "T-301" ? ; "METCBAR_MOLE" ; ; ICON2 ; ',
    '\\ PARAM NSTAGE = 8 ',
    '\\ \\ FEEDS FEED-SID = "S-112" FEED-STAGE = 1 /  FEED-SID = "S-115" FEED-STAGE = 7 '
    '/  FEED-SID = "S-111" FEED-STAGE = 8 ',
    '\\ \\ PRODUCTS PROD-STREAM = "S-114" PROD-STAGE = 1 PROD-PHASE = L2 P-S = N / '
    ' PROD-STREAM = "S-113" PROD-STAGE = 8 PROD-PHASE = L1 P-S = N ',
    '\\ \\ P-SPEC PRES-STAGE = 1 STAGE-PRES = 1.0 <20> <5> ',
    '\\ \\ "L1-COMPS" COMP1-LIST = ( H2O ) ',
    '\\ \\ "L2-COMPS" COMP2-LIST = ( TOL ) ',
    '\\ \\ T-EST TEMP-STAGE = 1 TEMP-EST = 40.0 <22> <4> \\ ',
])

NEW = {'R-101': R101, 'R-501': R501, 'F-501': F501, 'T-201': T201, 'T-301': T301}

t = open(BASE, encoding='utf-8', errors='ignore').read()
starts = [m.start() for m in re.finditer(r'\?\s*BLOCK\s+[A-Z0-9]+\s+"?[A-Za-z0-9-]+"?\s*\?', t)]
rep = []
for idx in range(len(starts) - 1, -1, -1):
    s0 = starts[idx]
    s1 = starts[idx + 1] if idx + 1 < len(starts) else len(t)
    m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"([^"]+)"\s*\?', t[s0:s1])
    if m and m.group(2) in NEW:
        t = t[:s0] + NEW[m.group(2)] + NL + t[s1:]
        rep.append(m.group(2))
print('已重写:', rep, flush=True)
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
print('Ready =', doc.Engine.Ready, flush=True)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 240:
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
    print('%3d| %s' % (i, x[:230]), flush=True)

print()
print('=== 结果抽查 ===', flush=True)
for b in ['M-101', 'E-101', 'R-101', 'T-201', 'T-301', 'T-401', 'T-402', 'T-403',
          'T-404', 'R-501', 'F-501', 'E-601']:
    n = doc.Tree.FindNode(r'\Data\Blocks\%s\Output' % b)
    if n is None:
        print('  %-6s 无输出节点' % b, flush=True); continue
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
    doc.SaveAs(r'D:\<化工工作区>\_probe\sim2_saved.bkp')
except Exception:
    pass
try:
    doc.Close()
except Exception:
    pass
print('DONE', flush=True)
