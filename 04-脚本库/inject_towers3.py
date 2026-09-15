# -*- coding: utf-8 -*-
"""按 Aspen 原格式（记录分隔符在行首、带换行）重写五个塔段"""
import time, re
import win32com.client as win32

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_模拟.bkp'
OUT = r'D:\<化工工作区>\_probe\full3.bkp'

NL = '\n'


def rf(bid, n, feeds, dist, bot, p1, df, rr, cond='TOTAL', reb='KETTLE', icon='FRACT1'):
    L = []
    L.append('? BLOCK RADFRAC "%s" ? ; "METCBAR_MOLE" ; ; %s ; ' % (bid, icon))
    L.append('\\ PARAM NSTAGE = %d OPT-PRES = "DP-COL" OPT-PRES-TOP = "DP-COND" NSTAGEMAX = %d '
             % (n, n + 1))
    L.append('\\ \\ "COL-CONFIG" CONDENSER = %s REBOILER = %s ' % (cond, reb))
    L.append('\\ \\ FEEDS '
             + ' / '.join('FEED-SID = "%s" FEED-STAGE = %d' % (s, st) for s, st in feeds)
             + ' ')
    L.append('\\ \\ PRODUCTS PROD-STREAM = "%s" PROD-STAGE = 1 PROD-PHASE = L P-S = N / '
             'PROD-STREAM = "%s" PROD-STAGE = %d PROD-PHASE = L P-S = N ' % (dist, bot, n))
    L.append('\\ \\ "P-SPEC2" PRES1 = %s <20> <10> ' % p1)
    L.append('\\ \\ "COL-SPECS" D:F = %s <-1> <0> BASIS-RDV = 0.0 <-1> <0> '
             'BASIS-RR = %s <-1> <0> ' % (df, rr))
    L.append('\\ \\ "KLL-VECS" ')
    L.append('\\ \\ "TRSZ-VECS" ')
    L.append('\\ \\ "PCKSR-VECS" \\ ')
    return NL.join(L)


T201 = NL.join([
    '? BLOCK RADFRAC "T-201" ? ; "METCBAR_MOLE" ; ; FRACT1 ; ',
    '\\ PARAM NSTAGE = 6 OPT-PRES = "DP-COL" OPT-PRES-TOP = "DP-COND" NSTAGEMAX = 7 ',
    '\\ \\ "COL-CONFIG" CONDENSER = NONE REBOILER = NONE ',
    '\\ \\ FEEDS FEED-SID = "S-107" FEED-STAGE = 1 / FEED-SID = "S-108" FEED-STAGE = 6 ',
    '\\ \\ PRODUCTS PROD-STREAM = "S-109" PROD-STAGE = 1 PROD-PHASE = V P-S = N / '
    'PROD-STREAM = "S-110" PROD-STAGE = 6 PROD-PHASE = L P-S = N ',
    '\\ \\ "P-SPEC2" PRES1 = 1.0 <20> <10> ',
    '\\ \\ "KLL-VECS" ',
    '\\ \\ "TRSZ-VECS" ',
    '\\ \\ "PCKSR-VECS" \\ ',
])

NEW = {
    'T-201': T201,
    'T-401': rf('T-401', 25, [('S-114', 12)], 'S-115', 'S-116', 1.0, 24.5, 2.0),
    'T-402': rf('T-402', 40, [('S-116', 20)], 'S-117', 'S-118', 1.0, 5.0, 3.0),
    'T-403': rf('T-403', 45, [('S-118', 22)], 'S-119', 'S-120', 0.6, 0.5, 18.0),
    'T-404': rf('T-404', 32, [('S-120', 18)], 'S-121', 'S-122', 0.3, 9.5, 3.0),
}

t = open(SRC, encoding='utf-8', errors='ignore').read()
starts = [m.start() for m in
          re.finditer(r'\?\s*BLOCK\s+[A-Z0-9]+\s+"?[A-Za-z0-9-]+"?\s*\?', t)]
rep = []
for idx in range(len(starts) - 1, -1, -1):
    s0 = starts[idx]
    s1 = starts[idx + 1] if idx + 1 < len(starts) else len(t)
    m = re.match(r'\?\s*BLOCK\s+RADFRAC\s+"([^"]+)"\s*\?', t[s0:s1])
    if m and m.group(1) in NEW:
        t = t[:s0] + NEW[m.group(1)] + NL + t[s1:]
        rep.append(m.group(1))
print('已替换:', rep, flush=True)
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(OUT)
time.sleep(4)
print('Ready =', doc.Engine.Ready, flush=True)
print()
for b in ['T-201', 'T-401', 'T-402', 'T-403', 'T-404']:
    n = doc.Tree.FindNode(r'\Data\Blocks\%s\Input' % b)
    hits = []
    for i in range(n.Elements.Count):
        ch = n.Elements.Item(i)
        if ch.Name in ('NSTAGE', 'PRES1', 'CONDENSER', 'REBOILER', 'FEED_STAGE',
                       'D:F', 'BASIS_RR', 'BASIS_RDV'):
            try:
                v = ch.Value
            except Exception:
                v = '?'
            hits.append('%s=%r' % (ch.Name, v))
    print('  %-7s %s' % (b, '; '.join(hits)), flush=True)
try:
    doc.Close()
except Exception:
    pass
