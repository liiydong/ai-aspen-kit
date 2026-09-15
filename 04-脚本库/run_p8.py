# -*- coding: utf-8 -*-
"""P8：标定 T-402/T-403/T-404 采出比，恢复 3-CP 产品收率"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_v20.bkp'
base = open(SRC, encoding='utf-8', errors='ignore').read()


def repl_block(t, bid, fn):
    starts = [s.start() for s in re.finditer(r'\?\s*BLOCK\s+', t)]
    for k in range(len(starts)):
        s0 = starts[k]
        s1 = starts[k + 1] if k + 1 < len(starts) else len(t)
        m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
        if m and m.group(2) == bid:
            return t[:s0] + fn(t[s0:s1]) + t[s1:]
    raise RuntimeError(bid)


def build(t402_df, t402_rr, t403_df, t404_df, out):
    t = base
    for bid, df, rr in [('T-402', t402_df, t402_rr), ('T-403', t403_df, None), ('T-404', t404_df, None)]:
        def fn(seg, df=df, rr=rr):
            seg = re.sub(r'D:F = [\d.]+ <-1> <0>', 'D:F = %.4f <-1> <0>' % df, seg, count=1)
            if rr:
                seg = re.sub(r'BASIS-RR = [\d.]+ <-1> <0>', 'BASIS-RR = %s <-1> <0>' % rr, seg, count=1)
            return seg
        t = repl_block(t, bid, fn)
    assert len(re.findall(r'\?\s*BLOCK\s+RADFRAC', t)) == 5
    open(out, 'w', encoding='utf-8', errors='ignore').write(t)


def run(tag, path):
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

    sev = [x for x in MSGS if 'ERROR' in x.upper()]
    print('%-30s 严重报错=%d' % (tag, len(sev)), flush=True)
    for b in ['T-401', 'T-402', 'T-403', 'T-404']:
        bb = r'\Data\Blocks\%s\Output' % b
        print('   %-6s RR=%-6s Ttop=%-9s Tbot=%-9s | 顶:3-CP=%-8.1f 4-CP=%-6.1f | 釜:3-CP=%-8.1f 4-CP=%-6.1f' % (
            b, g(bb + r'\MOLE_RR'), g(bb + r'\TOP_TEMP'), g(bb + r'\BOTTOM_TEMP'),
            c('S-115' if b == 'T-401' else ('S-117' if b == 'T-402' else ('S-119' if b == 'T-403' else 'S-121')), '3-CP'),
            c('S-115' if b == 'T-401' else ('S-117' if b == 'T-402' else ('S-119' if b == 'T-403' else 'S-121')), '4-CP'),
            c('S-116' if b == 'T-401' else ('S-118' if b == 'T-402' else ('S-120' if b == 'T-403' else 'S-122')), '3-CP'),
            c('S-116' if b == 'T-401' else ('S-118' if b == 'T-402' else ('S-120' if b == 'T-403' else 'S-122')), '4-CP')),
            flush=True)
    print('   >>> S-121 产品 = %.1f kg/h（3-CP %.1f，4-CP %.1f）；S-117 顶 = %.1f kg/h；S-119 顶 = %.1f kg/h' % (
        g(r'\Data\Streams\S-121\Output\MASSFLMX\MIXED'), c('S-121', '3-CP'), c('S-121', '4-CP'),
        g(r'\Data\Streams\S-117\Output\MASSFLMX\MIXED'), g(r'\Data\Streams\S-119\Output\MASSFLMX\MIXED')), flush=True)
    try:
        doc.SaveAs(path.replace('.bkp', '_s.bkp'))
    except Exception:
        pass
    try:
        doc.Close()
    except Exception:
        pass


cases = [
    ('T-402轻组分 D:F=0.010', 0.010, '5.0', 0.0050, 0.9500),
    ('T-402轻组分 D:F=0.006', 0.006, '5.0', 0.0050, 0.9500),
]
for tag, a, b, cc, d in cases:
    p = r'D:\<化工工作区>\_probe\p8_%s.bkp' % a
    build(a, b, cc, d, p)
    run(tag, p)
print('DONE')
