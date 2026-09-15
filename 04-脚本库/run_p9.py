# -*- coding: utf-8 -*-
"""P9：查 3-CP/4-CP 常压沸点 + 提高 T-404 回流比提纯"""
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


def build(t402_df, t403_df, t403_rr, t404_df, t404_rr, out):
    t = base
    for bid, df, rr in [('T-402', t402_df, None), ('T-403', t403_df, t403_rr), ('T-404', t404_df, t404_rr)]:
        def fn(seg, df=df, rr=rr):
            seg = re.sub(r'D:F = [\d.]+ <-1> <0>', 'D:F = %.4f <-1> <0>' % df, seg, count=1)
            if rr:
                seg = re.sub(r'BASIS-RR = [\d.]+ <-1> <0>', 'BASIS-RR = %s <-1> <0>' % rr, seg, count=1)
            return seg
        t = repl_block(t, bid, fn)
    open(out, 'w', encoding='utf-8', errors='ignore').write(t)


def run(tag, path, first=False):
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

    def g(p):
        n = doc.Tree.FindNode(p)
        if n is None:
            return None
        try:
            return n.Value
        except Exception:
            return None

    if first:
        print('=== 纯组分常压沸点校验 ===')
        for cid in ['3-CP', '4-CP', '3-MP', '4-MP', 'TOL', 'NAM', 'H2O']:
            v = g(r'\Data\Components\%s\Input\TB' % cid)
            if v is None:
                v = g(r'\Data\Components\Input\TB')
            print('   %-6s TB = %s' % (cid, v))
        print('   （若上面全是同一值，说明该路径不可用，改用 TB 列表）')
        # 试另一种路径
        n = doc.Tree.FindNode(r'\Data\Components\Input\TB')
        if n is not None:
            try:
                print('   TB 列表值个数:', n.Elements.Count)
            except Exception:
                pass

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

    m121 = g(r'\Data\Streams\S-121\Output\MASSFLMX\MIXED')
    cp = c('S-121', '3-CP')
    fp = c('S-121', '4-CP')
    pur = (cp / m121 * 100) if (m121 and m121 > 0) else 0
    print('%-26s S-121=%-9.1f 3-CP=%-9.1f 4-CP=%-6.2f 纯度=%.2f%% | S-119=%-6.1f S-117=%-6.1f' % (
        tag, m121 or 0, cp, fp, pur,
        g(r'\Data\Streams\S-119\Output\MASSFLMX\MIXED') or 0,
        g(r'\Data\Streams\S-117\Output\MASSFLMX\MIXED') or 0), flush=True)
    sev = [x for x in MSGS if 'SEVERE' in x or 'ERROR' in x.upper()]
    for x in sev[:3]:
        print('      |', x[:170], flush=True)
    try:
        doc.SaveAs(path.replace('.bkp', '_s.bkp'))
    except Exception:
        pass
    try:
        doc.Close()
    except Exception:
        pass


p = r'D:\<化工工作区>\_probe\p9a.bkp'
build(0.006, 0.0050, '18.0', 0.9500, '10.0', p)
run('T-404 RR=10', p, first=True)
p = r'D:\<化工工作区>\_probe\p9b.bkp'
build(0.006, 0.0050, '18.0', 0.9500, '25.0', p)
run('T-404 RR=25', p)
print('DONE')
