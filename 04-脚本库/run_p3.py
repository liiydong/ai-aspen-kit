# -*- coding: utf-8 -*-
"""P3：D:F=0.66（目标组成的甲苯摩尔分率），两点标定 S-111"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_v19.bkp'
base = open(SRC, encoding='utf-8', errors='ignore').read()


def mk(makeup, df, out):
    t = base
    m = re.search(r'\?\s*STREAM\s+MATERIAL\s+"S-111"\s*\?', t)
    i = m.start()
    nxt = re.search(r'\?\s*STREAM\s+MATERIAL\s+"S-1', t[i + 30:])
    j = i + 30 + nxt.start()
    seg = t[i:j]
    seg = re.sub(r'(TOTAL\s*=\s*)24\.47', r'\g<1>%s' % makeup, seg)
    seg = re.sub(r'(CID\s*=\s*TOL\s+FLOW\s*=\s*)24\.47', r'\g<1>%s' % makeup, seg)
    t = t[:i] + seg + t[j:]
    starts = [s.start() for s in re.finditer(r'\?\s*BLOCK\s+', t)]
    for k in range(len(starts) - 1, -1, -1):
        s0 = starts[k]
        s1 = starts[k + 1] if k + 1 < len(starts) else len(t)
        mm = re.match(r'\?\s*BLOCK\s+RADFRAC\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
        if mm and mm.group(1) == 'T-401':
            sg = t[s0:s1]
            sg = re.sub(r'D:F = [\d.]+ <-1> <0>', 'D:F = %.4f <-1> <0>' % df, sg, count=1)
            sg = re.sub(r'BASIS-RR = [\d.]+ <-1> <0>', 'BASIS-RR = 3.0 <-1> <0>', sg, count=1)
            t = t[:s0] + sg + t[s1:]
            break
    open(out, 'w', encoding='utf-8', errors='ignore').write(t)


def run(tag, path, dump_err=False):
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
    while time.time() - t0 < 300:
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
    print('%-22s Ttop=%-9s Tbot=%-9s | S-114 TOL=%-9.1f | S-116 TOL=%-8.1f 3-CP=%-9.1f | S-113 TOL=%-7.2f' % (
        tag, g(bb + r'\TOP_TEMP'), g(bb + r'\BOTTOM_TEMP'),
        c('S-114', 'TOL'), c('S-116', 'TOL'), c('S-116', '3-CP'), c('S-113', 'TOL')), flush=True)
    if dump_err:
        for x in MSGS:
            if 'ERROR' in x.upper():
                print('      |', x[:190], flush=True)
    try:
        doc.Close()
    except Exception:
        pass


print('目标 S-114 甲苯 = 2254.9 kg/h（瑞邦设计值 [1]）')
mk(0.10, 0.66, r'D:\<化工工作区>\_probe\p3a.bkp')
run('D:F=0.66 M=0.10', r'D:\<化工工作区>\_probe\p3a.bkp')
mk(0.20, 0.66, r'D:\<化工工作区>\_probe\p3b.bkp')
run('D:F=0.66 M=0.20', r'D:\<化工工作区>\_probe\p3b.bkp')
mk(0.14, 0.66, r'D:\<化工工作区>\_probe\p3c.bkp')
run('D:F=0.66 M=0.14', r'D:\<化工工作区>\_probe\p3c.bkp', dump_err=True)
print('DONE')
