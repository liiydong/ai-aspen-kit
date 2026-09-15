# -*- coding: utf-8 -*-
"""重写 T-401~T-404 为已验证可解析的原生格式，并运行"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_sim2.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_sim3.bkp'
NL = '\n'


def col(bid, nst, feed, fstage, p1, p2, rr, ttop, tbot, pi, po):
    return NL.join([
        '? BLOCK RADFRAC "%s" ? ; "METCBAR_MOLE" ; ; FRACT1 ; ' % bid,
        '\\ PARAM NSTAGE = %d NSTAGEMAX = %d ' % (nst, nst + 1),
        '\\ \\ PARAM2 ',
        '\\ \\ "COL-CONFIG" CONDENSER = TOTAL REBOILER = KETTLE ',
        '\\ \\ FEEDS FEED-SID = "%s" FEED-STAGE = %d ' % (feed, fstage),
        '\\ \\ PRODUCTS PROD-STREAM = "%s" PROD-STAGE = 1 PROD-PHASE = L P-S = N / '
        ' PROD-STREAM = "%s" PROD-STAGE = %d PROD-PHASE = L P-S = N ' % (p1, p2, nst),
        '\\ \\ "P-SPEC2" PRES1 = %s <20> <10> ' % pi,
        '\\ \\ "COL-SPECS" BASIS-RDV = 0.0 <-1> <0> BASIS-RR = %s <-1> <0> ' % rr,
        '\\ \\ T-EST TEMP-STAGE = 1 TEMP-EST = %s <22> <4> / '
        'TEMP-STAGE = %d TEMP-EST = %s <22> <4> ' % (ttop, nst, tbot),
        '\\ \\ "KLL-VECS" ',
        '\\ \\ "TRSZ-VECS" ',
        '\\ \\ "PCKSR-VECS" \\ ',
    ])


NEW = {
    'T-401': col('T-401', 25, 'S-114', 12, 'S-115', 'S-116', '2.0', '60.0', '100.0', '1.0', '1.0'),
    'T-402': col('T-402', 40, 'S-116', 20, 'S-117', 'S-118', '3.0', '80.0', '105.0', '1.0', '1.0'),
    'T-403': col('T-403', 45, 'S-118', 22, 'S-119', 'S-120', '18.0', '90.0', '110.0', '0.6', '0.6'),
    'T-404': col('T-404', 32, 'S-120', 18, 'S-121', 'S-122', '3.0', '120.0', '150.0', '0.3', '0.3'),
}

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
    print('%3d| %s' % (i, x[:230]), flush=True)

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
    doc.SaveAs(r'D:\<化工工作区>\_probe\sim3_saved.bkp')
except Exception:
    pass
try:
    doc.Close()
except Exception:
    pass
print('DONE', flush=True)
