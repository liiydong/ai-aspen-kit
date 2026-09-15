# -*- coding: utf-8 -*-
"""Q1：核对组分身份 + 让 T-403 真正把 4-CP 赶到塔顶"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_最终版.bkp'
t = open(SRC, encoding='utf-8', errors='ignore').read()
print('重复性: RADFRAC=%d T-401=%d S-111=%d' % (
    len(re.findall(r'\?\s*BLOCK\s+RADFRAC', t)),
    len(re.findall(r'BLOCK\s+RADFRAC\s*\n?\s*"T-401"', t)),
    len(re.findall(r'\?\s*STREAM\s+MATERIAL\s*\n?\s*"S-111"', t))))

ci = t.find('? COMPONENTS MAIN ?')
cj = t.find('? COMPONENTS "COMP-LIST"', ci)
print()
print('=== 组分清单（CID / DBNAME1 / ANAME1）===')
for m in re.finditer(r'CID = "?([A-Za-z0-9-]+)"?\s+ANAME = ([A-Z0-9-]+)\s+OUTNAME = "?([A-Za-z0-9-]+)"?\s+TYPE = (\w+)\s+DBNAME1 = "([^"]+)"\s+ANAME1 = "([^"]+)"', t[ci:cj]):
    print('  %-8s DB=%-22s ANAME1=%-14s' % (m.group(1), m.group(5), m.group(6)))
print()


def repl_block(x, bid, fn):
    starts = [s.start() for s in re.finditer(r'\?\s*BLOCK\s+', x)]
    for k in range(len(starts)):
        s0 = starts[k]
        s1 = starts[k + 1] if k + 1 < len(starts) else len(x)
        m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', x[s0:s1])
        if m and m.group(2) == bid:
            return x[:s0] + fn(x[s0:s1]) + x[s1:]
    raise RuntimeError(bid)


def build(df403, rr403, out):
    x = t
    def fn(seg):
        seg = re.sub(r'D:F = [\d.]+ <-1> <0>', 'D:F = %.4f <-1> <0>' % df403, seg, count=1)
        seg = re.sub(r'BASIS-RR = [\d.]+ <-1> <0>', 'BASIS-RR = %s <-1> <0>' % rr403, seg, count=1)
        return seg
    x = repl_block(x, 'T-403', fn)
    open(out, 'w', encoding='utf-8', errors='ignore').write(x)


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

    def comp(s, cid):
        n = doc.Tree.FindNode(r'\Data\Streams\%s\Output\MASSFLOW3' % s)
        if n is None:
            return None
        try:
            for i in range(n.Elements.Count):
                e = n.Elements.Item(i)
                if e.Name.upper() == cid.upper():
                    return float(e.Value) if e.Value else 0.0
        except Exception:
            pass
        return None

    print('=' * 78)
    print(tag)
    for b, tp, bt in [('T-403', 'S-119', 'S-120')]:
        bb = r'\Data\Blocks\%s\Output' % b
        print('  %s RR=%s Ttop=%s Tbot=%s' % (b, g(bb + r'\MOLE_RR'), g(bb + r'\TOP_TEMP'), g(bb + r'\BOTTOM_TEMP')))
        print('     %s 总=%s | 3-CP=%s 4-CP=%s' % (tp, g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % tp),
                                                   comp(tp, '3-CP'), comp(tp, '4-CP')))
        print('     %s 总=%s | 3-CP=%s 4-CP=%s' % (bt, g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % bt),
                                                   comp(bt, '3-CP'), comp(bt, '4-CP')))
    print('     S-121 总=%s | 3-CP=%s 4-CP=%s' % (g(r'\Data\Streams\S-121\Output\MASSFLMX\MIXED'),
                                                  comp('S-121', '3-CP'), comp('S-121', '4-CP')))
    sev = [x for x in MSGS if 'Summary' in x or 'Errors' in x or 'Terminal' in x or 'Severe' in x]
    for x in sev:
        print('     |', x[:150])
    try:
        doc.Close()
    except Exception:
        pass


build(0.0080, '20.0', r'D:\<化工工作区>\_probe\q1a.bkp')
run('T-403 D:F=0.0080 RR=20', r'D:\<化工工作区>\_probe\q1a.bkp')
build(0.0120, '25.0', r'D:\<化工工作区>\_probe\q1b.bkp')
run('T-403 D:F=0.0120 RR=25', r'D:\<化工工作区>\_probe\q1b.bkp')
print('DONE')
