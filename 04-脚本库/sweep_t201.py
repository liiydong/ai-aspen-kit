# -*- coding: utf-8 -*-
"""T-201 变体扫描：用控制面板定位底部气相进料报错"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_sim6.bkp'
T0 = open(BASE, encoding='utf-8', errors='ignore').read()
NL = '\n'


def sec(v):
    if v == 1:      # SUM-RATES
        p = '\\ PARAM NSTAGE = 6 ALGORITHM = "SUM-RATES" INIT-OPTION = STANDARD NSTAGEMAX = 7 '
        feeds = ('\\ \\ FEEDS FEED-SID = "S-107" FEED-STAGE = 1 /  '
                 'FEED-SID = "S-108" FEED-STAGE = 6 ')
        extra = []
    elif v == 2:    # 只保留气相进料
        p = '\\ PARAM NSTAGE = 6 NSTAGEMAX = 7 '
        feeds = '\\ \\ FEEDS FEED-SID = "S-108" FEED-STAGE = 6 '
        extra = []
    else:           # 完全照 pyrolysis absorber
        p = ('\\ PARAM NSTAGE = 6 ALGORITHM = STANDARD INIT-OPTION = STANDARD '
             'MAXOL = 100 NO-PHASE = 3 OPT-PRES-TOP = "DP-COND" '
             'CONV-METH = STANDARD NSTAGEMAX = 7 ')
        feeds = ('\\ \\ FEEDS FEED-SID = "S-107" FEED-STAGE = 1 '
                 'FEED-CONVE2 = "ON-STAGE" /  FEED-SID = "S-108" FEED-STAGE = 6 '
                 'FEED-CONVEN = "ON-STAGE" FEED-CONVE2 = "ON-STAGE" ')
        extra = ['\\ \\ "COL-SPECS" DP-STAGE = 0.0 <75> <5> BASIS-RDV = 1.0 <0> <0> '
                 'DP-COND = 0.0 <75> <5> ']
    L = ['? BLOCK RADFRAC "T-201" ? ; "METCBAR_MOLE" ; ; ABSBR1 ; ',
         p,
         '\\ \\ PARAM2 ',
         '\\ \\ "COL-CONFIG" CONDENSER = NONE REBOILER = NONE ',
         feeds,
         '\\ \\ PRODUCTS PROD-STREAM = "S-109" PROD-STAGE = 1 PROD-PHASE = V P-S = N / '
         ' PROD-STREAM = "S-110" PROD-STAGE = 6 PROD-PHASE = L P-S = N ',
         '\\ \\ "P-SPEC2" PRES1 = 1.0 <20> <5> ']
    L += extra
    if v != 3:
        L.append('\\ \\ "COL-SPECS" BASIS-RDV = 1.0 <0> <0> ')
    L += ['\\ \\ T-EST TEMP-STAGE = 1 TEMP-EST = 40.0 <22> <4> / '
          'TEMP-STAGE = 6 TEMP-EST = 80.0 <22> <4> ',
          '\\ \\ "KLL-VECS" ',
          '\\ \\ "TRSZ-VECS" ',
          '\\ \\ "PCKSR-VECS" \\ ']
    return NL.join(L)


def run(v):
    m = re.search(r'\?\s*BLOCK\s+RADFRAC\s+"?T-201\s*"?\s*\?', T0)
    nxt = re.search(r'\n\?\s*BLOCK\s', T0[m.start() + 10:])
    end = m.start() + 10 + nxt.start() if nxt else len(T0)
    txt = T0[:m.start()] + sec(v) + NL + T0[end:]
    out = r'D:\<化工工作区>\_probe\t201_%d.bkp' % v
    open(out, 'w', encoding='utf-8', errors='ignore').write(txt)

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
    doc.InitFromArchive2(out)
    time.sleep(3)
    doc.Engine.Run2(False)
    t1 = time.time()
    while time.time() - t1 < 200:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.25)
        if time.time() - t1 > 8:
            try:
                if not bool(doc.Engine.IsRunning):
                    break
            except Exception:
                break
    for _ in range(40):
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)

    rel = [x for x in msgs if 'T-201' in x or 'BOTTOM STAGE' in x or 'SEVERE' in x
           or 'REST OF BLOCK' in x]
    print('变体 %d -> 面板 %d 条，T-201 相关 %d 条' % (v, len(msgs), len(rel)), flush=True)
    for x in rel[:10]:
        print('    |', x[:200], flush=True)
    if not rel:
        print('    ✓ T-201 无报错')
    nn = doc.Tree.FindNode(r'\Data\Blocks\T-201\Output')
    if nn is not None:
        got = []
        for f in ['TEMP', 'PRES', 'DUTY', 'B-PRES']:
            try:
                e = nn.Elements.Item(f)
                if e is not None:
                    got.append('%s=%s' % (f, e.Value))
            except Exception:
                pass
        print('    T-201 结果:', ', '.join(got), flush=True)
    try:
        doc.Close()
    except Exception:
        pass
    time.sleep(1)


for v in (1, 2, 3):
    try:
        run(v)
    except Exception as ex:
        print('变体 %d 异常 %s' % (v, ex))
    print()
print('ALL DONE')
