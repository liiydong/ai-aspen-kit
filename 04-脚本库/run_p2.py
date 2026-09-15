# -*- coding: utf-8 -*-
"""P2：T-401 D:F 按真实甲苯摩尔分率 0.50；标定 S-111 使甲苯循环量≈2254.9"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_v19.bkp'
base = open(SRC, encoding='utf-8', errors='ignore').read()


def mk(makeup, df, out):
    t = base
    # S-111 流量
    m = re.search(r'\?\s*STREAM\s+MATERIAL\s+"S-111"\s*\?', t)
    i = m.start()
    nxt = re.search(r'\?\s*STREAM\s+MATERIAL\s+"S-1', t[i + 30:])
    j = i + 30 + nxt.start()
    seg = t[i:j]
    seg = re.sub(r'(TOTAL\s*=\s*)24\.47', r'\g<1>%s' % makeup, seg)
    seg = re.sub(r'(CID\s*=\s*TOL\s+FLOW\s*=\s*)24\.47', r'\g<1>%s' % makeup, seg)
    t = t[:i] + seg + t[j:]
    # T-401
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
    print('%-24s T-401 Ttop=%-9s Tbot=%-9s | S-114 TOL=%-9.1f | S-115 TOL=%-9.1f | S-116 TOL=%-8.1f | S-116 3-CP=%-9.1f' % (
        tag, g(bb + r'\TOP_TEMP'), g(bb + r'\BOTTOM_TEMP'),
        c('S-114', 'TOL'), c('S-115', 'TOL'), c('S-116', 'TOL'), c('S-116', '3-CP')), flush=True)
    sev = [x for x in MSGS if re.search(r'ERROR', x)]
    for x in sev[:4]:
        print('      |', x[:150], flush=True)
    try:
        doc.Close()
    except Exception:
        pass


print('目标：S-114 甲苯 ≈ 2254.9 kg/h（瑞邦设计值）')
for mk_, df_ in [(0.50, 0.50), (0.95, 0.50), (0.95, 0.60)]:
    p = r'D:\<化工工作区>\_probe\p2_%s_%s.bkp' % (str(mk_).replace('.', ''), str(df_).replace('.', ''))
    mk(mk_, df_, p)
    run('M=%.2f D:F=%.2f kmol/h' % (mk_, df_), p)
print('DONE')
