# -*- coding: utf-8 -*-
"""改用单反斜杠分隔符重新注入塔规格"""
import time, re
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\tp_ok.bkp'
OUT = r'D:\<化工工作区>\_probe\full2.bkp'
SEP = ' \\ '          # 单个反斜杠作记录分隔


def rf(bid, n, feeds, dist, bot, p1, df, rr, cond='TOTAL', reb='KETTLE'):
    feedtxt = ' / '.join('FEED-SID = "%s" FEED-STAGE = %d' % (s, st) for s, st in feeds)
    return ('? BLOCK RADFRAC "%s" ? ; "METCBAR_MOLE" ; ; FRACT1 ; '
            '%sPARAM NSTAGE = %d OPT-PRES = "DP-COL" OPT-PRES-TOP = "DP-COND" NSTAGEMAX = %d '
            '%s"COL-CONFIG" CONDENSER = %s REBOILER = %s '
            '%sFEEDS %s '
            '%sPRODUCTS PROD-STREAM = "%s" PROD-STAGE = 1 PROD-PHASE = L P-S = N / '
            'PROD-STREAM = "%s" PROD-STAGE = %d PROD-PHASE = L P-S = N '
            '%s"P-SPEC2" PRES1 = %s <20> <10> '
            '%s"COL-SPECS" D:F = %s <-1> <0> BASIS-RDV = 0.0 <-1> <0> BASIS-RR = %s <-1> <0> '
            '%s"KLL-VECS" %s"TRSZ-VECS" %s"PCKSR-VECS" \\ '
            % (bid, SEP, n, n + 1, SEP, cond, reb, SEP, feedtxt, SEP,
               dist, bot, n, SEP, p1, SEP, df, rr, SEP, SEP, SEP))


T201 = ('? BLOCK RADFRAC "T-201" ? ; "METCBAR_MOLE" ; ; FRACT1 ; '
        '%sPARAM NSTAGE = 6 OPT-PRES = "DP-COL" OPT-PRES-TOP = "DP-COND" NSTAGEMAX = 7 '
        '%s"COL-CONFIG" CONDENSER = NONE REBOILER = NONE '
        '%sFEEDS FEED-SID = "S-107" FEED-STAGE = 1 / FEED-SID = "S-108" FEED-STAGE = 6 '
        '%sPRODUCTS PROD-STREAM = "S-109" PROD-STAGE = 1 PROD-PHASE = V P-S = N / '
        'PROD-STREAM = "S-110" PROD-STAGE = 6 PROD-PHASE = L P-S = N '
        '%s"P-SPEC2" PRES1 = 1.0 <20> <10> '
        '%s"KLL-VECS" %s"TRSZ-VECS" %s"PCKSR-VECS" \\ '
        % (SEP, SEP, SEP, SEP, SEP, SEP, SEP, SEP))

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
        t = t[:s0] + NEW[m.group(1)] + ' \n' + t[s1:]
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
print()
print('=== 读回 ===', flush=True)
for b in ['T-201', 'T-401', 'T-402', 'T-403', 'T-404']:
    row = []
    for f in ['NSTAGE', 'PRES1']:
        n = doc.Tree.FindNode(r'\Data\Blocks\%s\Input\%s' % (b, f))
        row.append('%s=%s' % (f, n.Value if n is not None else 'None'))
    print('  %-7s %s' % (b, '  '.join(row)), flush=True)
try:
    doc.Engine.Run2()
    print('Run2 完成', flush=True)
except Exception as e:
    print('Run2 抛出', str(e)[:120], flush=True)
print()
for s in ['S-106', 'S-201', 'S-301', 'S-401']:
    n = doc.Tree.FindNode(r'\Data\Streams\%s\Output\TEMP' % s)
    print('  %s.Out.TEMP = %r' % (s, n.Value if n is not None else None), flush=True)
try:
    doc.SaveAs(r'D:\<化工工作区>\_probe\full2_saved.bkp')
except Exception:
    pass
try:
    doc.Close()
except Exception:
    pass
