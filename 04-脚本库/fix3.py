# -*- coding: utf-8 -*-
"""按 Aspen 原始风格重写 3 个反应器 + T-201 + T-301"""
import time, re
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\fix2.bkp'
OUT = r'D:\<化工工作区>\_probe\fix3.bkp'
NL = '\n'


def k(n, base):
    """复刻 Aspen 的后缀规则：前两条裸关键字，第 k(>=3) 条用 {base}{k-2}"""
    return base if n <= 2 else '%s%d' % (base, n - 2)


def rstoic(bid, temp, pres, stoic, convex, prod):
    L = ['? BLOCK RSTOIC "%s" ? ; "METCBAR_MOLE" ; ; ICON1 ; ' % bid,
         '\\ PARAM TEMP = %s <22> <4> PRES = %s <20> <5> SPEC-OPT = TP ' % (temp, pres)]
    for i, (rn, cid, co) in enumerate(stoic, 1):
        s = '' if i <= 2 else str(i - 2)
        L.append('\\ \\ STOIC%s REACNO%s = %d STOIC-CID%s = "%s" STOIC-SSID%s = MIXED '
                 'COEF%s = %s <0> <0> ' % (s, s, rn, s, cid, s, s, co))
    for i, (rn, key, cv) in enumerate(convex, 1):
        s = '' if i <= 2 else str(i - 2)
        L.append('\\ \\ CONVEX%s EXT-REACNO%s = %d KEY-SSID%s = MIXED KEY-CID%s = "%s" '
                 'CONV%s = %s <0> <0> ' % (s, s, rn, s, s, key, s, cv))
    L.append('\\ \\ PRODUCTS SID = "%s" \\ ' % prod)
    return NL.join(L)


R101 = rstoic('R-101', '405.0', '1.8',
              [(1, '3-MP', -1.0), (1, 'NH3', -1.0), (1, 'O2', -1.5),
               (1, '3-CP', 1.0), (1, 'H2O', 3.0),
               (2, '3-MP', -1.0), (2, 'O2', -7.5), (2, 'CO2', 6.0),
               (2, 'H2O', 3.0), (2, 'HCN', 1.0)],
              [(1, '3-MP', '.86'), (2, '3-MP', '.02')], 'S-106')

R501 = rstoic('R-501', '140.0', '4.0',
              [(1, '3-CP', -1.0), (1, 'H2O', -1.0), (1, 'NAM', 1.0),
               (2, 'NAM', -1.0), (2, 'H2O', -1.0), (2, 'NAC', 1.0), (2, 'NH3', 1.0)],
              [(1, '3-CP', '.985'), (2, 'NAM', '.005')], 'S-201')

F501 = rstoic('F-501', '60.0', '1.0',
              [(1, 'NAOH', -2.0), (1, 'H2SO4', -1.0), (1, 'NA2SO4', 1.0), (1, 'H2O', 2.0)],
              [(1, 'H2SO4', '1.0')], 'S-301')

# T-201 吸收塔：补两个进料板 + 一个采出规格
T201 = NL.join([
    '? BLOCK RADFRAC "T-201" ? ; "METCBAR_MOLE" ; ; FRACT1 ; ',
    '\\ PARAM NSTAGE = 6 OPT-PRES = "DP-COL" OPT-PRES-TOP = "DP-COND" NSTAGEMAX = 7 ',
    '\\ \\ "COL-CONFIG" CONDENSER = NONE REBOILER = NONE ',
    '\\ \\ FEEDS FEED-SID = "S-107" FEED-STAGE = 1 / FEED-SID = "S-108" FEED-STAGE = 6 ',
    '\\ \\ PRODUCTS PROD-STREAM = "S-109" PROD-STAGE = 1 PROD-PHASE = V P-S = N / '
    'PROD-STREAM = "S-110" PROD-STAGE = 6 PROD-PHASE = L P-S = N ',
    '\\ \\ "P-SPEC2" PRES1 = 1.0 <20> <10> ',
    '\\ \\ "COL-SPECS" D:F = 175.0 <-1> <0> ',
    '\\ \\ "KLL-VECS" ',
    '\\ \\ "TRSZ-VECS" ',
    '\\ \\ "PCKSR-VECS" \\ ',
])

# T-301 萃取塔：照 Biodiesel 示例的 EXTRACT 格式
T301 = NL.join([
    '? BLOCK EXTRACT "T-301" ? ; "METCBAR_MOLE" ; ; ICON2 ; ',
    '\\ PARAM NSTAGE = 8 ',
    '\\ \\ FEEDS FEED-SID = "S-112" FEED-STAGE = 1 / FEED-SID = "S-115" FEED-STAGE = 8 ',
    '\\ \\ PRODUCTS PROD-STREAM = "S-114" PROD-STAGE = 1 PROD-PHASE = L2 P-S = N / '
    'PROD-STREAM = "S-113" PROD-STAGE = 8 PROD-PHASE = L1 P-S = N ',
    '\\ \\ P-SPEC PRES-STAGE = 1 STAGE-PRES = 1.0 <20> <5> ',
    '\\ \\ "L1-COMPS" COMP1-LIST = ( H2O ) ',
    '\\ \\ "L2-COMPS" COMP2-LIST = ( TOL ) ',
    '\\ \\ T-EST TEMP-STAGE = 1 TEMP-EST = 40.0 <22> <4> \\ ',
])

t = open(SRC, encoding='utf-8', errors='ignore').read()
starts = [m.start() for m in
          re.finditer(r'\?\s*BLOCK\s+[A-Z0-9]+\s+"?[A-Za-z0-9-]+"?\s*\?', t)]
NEW = {'R-101': R101, 'R-501': R501, 'F-501': F501,
       'T-201': T201, 'T-301': T301}
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

# 运行并抓控制面板
import pythoncom
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
win32.WithEvents(doc, Sink)
doc.InitFromArchive2(OUT)
time.sleep(4)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 220:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 10:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(40):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)
print()
print('=== 控制面板 (%d 条) ===' % len(MSGS), flush=True)
for m in MSGS:
    print('   ', m[:300], flush=True)
print()
print('=== 结果抽查 ===', flush=True)
for b in ['M-101', 'E-101', 'R-101', 'T-201', 'T-301', 'T-401', 'R-501', 'F-501', 'E-601']:
    n = doc.Tree.FindNode(r'\Data\Blocks\%s\Output' % b)
    if n is None:
        print('  %-7s 无' % b, flush=True)
        continue
    vals = []
    for f in ['TEMP', 'PRES', 'DUTY', 'MOLE_RR', 'B-PRES']:
        try:
            x = n.Elements.Item(f)
            if x is not None and x.Value not in (None, ''):
                vals.append('%s=%s' % (f, x.Value))
        except Exception:
            pass
    print('  %-7s %s' % (b, ','.join(vals) if vals else '(空)'), flush=True)
try:
    doc.SaveAs(r'D:\<化工工作区>\_probe\fix3_saved.bkp')
except Exception:
    pass
try:
    doc.Close()
except Exception:
    pass
