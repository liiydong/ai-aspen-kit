# -*- coding: utf-8 -*-
"""P5：整段重建 Sep（含惰性气体）+ T-401 D:F 反解甲苯循环量"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_v19.bkp'
base = open(SRC, encoding='utf-8', errors='ignore').read()


def wr(s, w=74):
    out, line = [], ''
    for tok in s.split(' '):
        if len(line) + len(tok) + 1 > w:
            out.append(line)
            line = tok
        else:
            line = (line + ' ' + tok) if line else tok
    if line:
        out.append(line)
    return '\n'.join(out)


FR = {'H2O': 0.999, 'TOL': 0.0005, '3-CP': 0.005, '4-CP': 0.005,
      '3-MP': 0.005, '4-MP': 0.005, 'HCN': 0.50, 'NH3': 0.99,
      'O2': 0.99, 'N2': 0.99, 'CO2': 0.99, 'CO': 0.99,
      'NAM': 0.999, 'NAC': 0.999, 'AIR': 0.999, 'NAOH': 0.999,
      'H2SO4': 0.999, 'NA2SO4': 0.999}


def mk(makeup, df, out):
    t = base
    # 1) S-111
    m = re.search(r'\?\s*STREAM\s+MATERIAL\s+"S-111"\s*\?', t)
    i = m.start()
    nxt = re.search(r'\?\s*STREAM\s+MATERIAL\s+"S-1', t[i + 30:])
    j = i + 30 + nxt.start()
    seg = t[i:j]
    seg = re.sub(r'(TOTAL\s*=\s*)24\.47', r'\g<1>%s' % makeup, seg)
    seg = re.sub(r'(CID\s*=\s*TOL\s+FLOW\s*=\s*)24\.47', r'\g<1>%s' % makeup, seg)
    t = t[:i] + seg + t[j:]

    # 2) 重建 Sep 段
    ci = t.find('? COMPONENTS MAIN ?')
    cj = t.find('? COMPONENTS "COMP-LIST"', ci)
    cids = re.findall(r'CID = "?([A-Za-z0-9-]+)"?', t[ci:cj])
    recs = ['PARAM-STREAM = "S-113" SUBSTREAM = MIXED COMPS = "%s" FRACS = %s <0> <0> ' % (c, FR[c])
            for c in cids if c in FR]
    para = ('? BLOCK SEP "T-301" ? ; "METCBAR_MOLE" ; ; ICON2 ; \\ '
            'DESCRIPTION DESCRIPTION = "Toluene extraction 99.5 pct per Ruibang" \\ \\ '
            'PARAM1 PRES1 = 1.0 <20> <5> \\ \\ PARAM ' + ' /  '.join(recs) + ' \\ ')
    starts = [s.start() for s in re.finditer(r'\?\s*BLOCK\s+', t)]
    loc = None
    for k in range(len(starts) - 1, -1, -1):
        s0 = starts[k]
        s1 = starts[k + 1] if k + 1 < len(starts) else len(t)
        mm = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
        if mm and mm.group(2) == 'T-301':
            loc = (s0, s1)
            break
    t = t[:loc[0]] + wr(para) + t[loc[1]:]

    # 3) T-401
    starts = [s.start() for s in re.finditer(r'\?\s*BLOCK\s+', t)]
    for kk in range(len(starts) - 1, -1, -1):
        s0 = starts[kk]
        s1 = starts[kk + 1] if kk + 1 < len(starts) else len(t)
        mm = re.match(r'\?\s*BLOCK\s+RADFRAC\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
        if mm and mm.group(1) == 'T-401':
            sg = t[s0:s1]
            sg = re.sub(r'D:F = [\d.]+ <-1> <0>', 'D:F = %.4f <-1> <0>' % df, sg, count=1)
            sg = re.sub(r'BASIS-RR = [\d.]+ <-1> <0>', 'BASIS-RR = 3.0 <-1> <0>', sg, count=1)
            t = t[:s0] + sg + t[s1:]
            break
    open(out, 'w', encoding='utf-8', errors='ignore').write(t)


def run(tag, path, err=False):
    MSGS = []

    class Sink:
        def OnControlPanelMessage(self, *a):
            s = ' '.join(str(x) for x in a).strip()
            if s and s != 'False':
                MSGS.append(s)

    doc = win32.DispatchEx('Apwn.Document')
    doc.SuppressDialogs = True
    try:
        win32.WithEvents(doc, Sink)
    except Exception:
        pass
    doc.InitFromArchive2(path)
    time.sleep(3)
    doc.Engine.Run2(False)
    t0 = time.time()
    while time.time() - t0 < 320:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.25)
        if time.time() - t0 > 10:
            try:
                if not bool(doc.Engine.IsRunning):
                    break
            except Exception:
                break
    for _ in range(50):
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)

    def g(p):
        n = doc.Tree.FindNode(p)
        if n is None:
            return None
        try:
            return n.Value
        except Exception:
            return None

    def c(s, cid):
        n = doc.Tree.FindNode(r'\Data\Streams\%s\Output\MASSFLOW3' % s)
        if n is None:
            return 0.0
        try:
            for i in range(n.Elements.Count):
                e = n.Elements.Item(i)
                if e.Name.upper() == cid.upper():
                    return float(e.Value) if e.Value else 0.0
        except Exception:
            pass
        return 0.0

    bb = r'\Data\Blocks\T-401\Output'
    print('%-24s Ttop=%-9s Tbot=%-9s | S-114 TOL=%-9.1f | S-115 TOL=%-9.1f | S-116 TOL=%-7.1f 3-CP=%-9.1f' % (
        tag, g(bb + r'\TOP_TEMP'), g(bb + r'\BOTTOM_TEMP'),
        c('S-114', 'TOL'), c('S-115', 'TOL'), c('S-116', 'TOL'), c('S-116', '3-CP')), flush=True)
    if err:
        for x in MSGS:
            if 'ERROR' in x.upper() or 'HCALC' in x:
                print('      |', x[:185], flush=True)
    try:
        doc.Close()
    except Exception:
        pass


print('目标 S-114 甲苯 = 2254.9 kg/h')
mk(0.01, 0.659, r'D:\<化工工作区>\_probe\p5a.bkp')
run('M=0.01 D:F=0.659', r'D:\<化工工作区>\_probe\p5a.bkp', err=True)
mk(0.01, 0.700, r'D:\<化工工作区>\_probe\p5b.bkp')
run('M=0.01 D:F=0.700', r'D:\<化工工作区>\_probe\p5b.bkp')
print('DONE')
