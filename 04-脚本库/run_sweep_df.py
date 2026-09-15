# -*- coding: utf-8 -*-
"""扫描循环甲苯量（T-401 D:F）与萃取级数，判断 3-CP 回收率的上限"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\_probe\bipA2.bkp'
base = open(SRC, encoding='utf-8', errors='ignore').read()


def make(df, nstage301, out):
    t = base
    # T-401 D:F
    t = re.sub(r'"COL-SPECS" D:F = [\d.]+ <-1> <0>', '"COL-SPECS" D:F = %.4f <-1> <0>' % df, t, count=1)
    # T-301 NSTAGE（只改 T-301 段里的 PARAM NSTAGE）
    if nstage301 != 8:
        i = t.find('BLOCK EXTRACT "T-301"')
        j = t.find('BLOCK ', i + 10)
        seg = t[i:j]
        seg2 = re.sub(r'PARAM NSTAGE = \d+', 'PARAM NSTAGE = %d' % nstage301, seg, count=1)
        t = t[:i] + seg2 + t[j:]
    open(out, 'w', encoding='utf-8', errors='ignore').write(t)


def run(tag, path):
    doc = win32.DispatchEx('Apwn.Document')
    try:
        doc.SuppressDialogs = True
    except Exception:
        pass
    doc.InitFromArchive2(path)
    time.sleep(3)
    doc.Engine.Run2(False)
    t0 = time.time()
    while time.time() - t0 < 240:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.25)
        if time.time() - t0 > 10:
            try:
                if not bool(doc.Engine.IsRunning):
                    break
            except Exception:
                break
    for _ in range(40):
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
        b = r'\Data\Streams\%s\Output' % s
        n = doc.Tree.FindNode(b + r'\MASSFLOW3')
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

    cp114 = comp('S-114', '3-CP')
    cp113 = comp('S-113', '3-CP')
    tol114 = comp('S-114', 'TOL')
    tol113 = comp('S-113', 'TOL')
    h2o114 = comp('S-114', 'H2O')
    tot = None
    if cp114 is not None and cp113 is not None:
        tot = cp114 + cp113
    rec = (cp114 / 1266.98 * 100) if cp114 is not None else None
    print('%-26s S-114:3-CP=%-10s TOL=%-10s H2O=%-8s | S-113:3-CP=%-10s TOL=%-8s | 回收率=%.1f%%' % (
        tag,
        ('%.2f' % cp114) if cp114 is not None else 'None',
        ('%.1f' % tol114) if tol114 is not None else 'None',
        ('%.2f' % h2o114) if h2o114 is not None else 'None',
        ('%.2f' % cp113) if cp113 is not None else 'None',
        ('%.2f' % tol113) if tol113 is not None else 'None',
        rec if rec is not None else -1), flush=True)
    try:
        doc.Close()
    except Exception:
        pass


print('%-26s %s' % ('条件', '结果'), flush=True)
for df in [0.70, 0.85, 0.92]:
    o = r'D:\<化工工作区>\_probe\sw_df%02d.bkp' % int(df * 100)
    make(df, 8, o)
    run('T-401 D:F=%.2f 级数8' % df, o)
o = r'D:\<化工工作区>\_probe\sw_n16.bkp'
make(0.92, 16, o)
run('T-401 D:F=0.92 级数16', o)
print('DONE')
