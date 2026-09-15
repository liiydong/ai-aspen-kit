# -*- coding: utf-8 -*-
"""在干净模型（无损坏结构）上注入真实二元参数，并与基线对照"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_s13.bkp'
OUT = r'D:\<化工工作区>\_probe'
print('基线文件存在:', os.path.exists(BASE), os.path.getsize(BASE) if os.path.exists(BASE) else '')

PAIRS = [
    ('O2', 'H2O', [-3.28368021, 17.8247855, 1617.10561, -381.901258, 0.3,
                   0, 0, 0, 0, 0, 0, 1000]),
    ('H2O', 'CO2', [10.0640, 10.0640, -3268.1350, -3268.1350, .20]),
    ('H2O', 'HCN', [.0, .0, 505.50, .0, .30]),
    ('H2O', 'NH3', [-.5440720, -.16424220, 1678.4690, -1027.5250, .20]),
    ('H2O', 'TOL', [627.05280, -247.87920, -27269.360, 14759.760, .20,
                    0.0, -92.71820, 35.5820, 0.0, 0.0, 264.150, 366.150]),
]


def fmt(v):
    s = ('%.8f' % v).rstrip('0').rstrip('.')
    return s if s not in ('', '-0') else '0.0'


def rec(c1, c2, vals):
    parts = ' '.join('UVAL%d = %s <0> <0>' % (k + 1, fmt(v)) for k, v in enumerate(vals))
    return ('\n\\ BPVAL PARAMNAME2 = NRTL CID1 = %s CID2 = %s UNITROW2 = 0 '
            'TUNITROW2 = 22 TUNITLABEL2 = K %s \\' % (c1, c2, parts))


t0 = open(BASE, encoding='utf-8', errors='ignore').read()
print('MOLEC-STRUCT 出现次数:', t0.count('BONDS'))

m = re.search(r'(ESTIMATE = NO\s*\\)', t0)
print('锚点:', bool(m))
newrecs = ''.join(rec(*p) for p in PAIRS)
t = t0[:m.end()] + newrecs + t0[m.end():]

p = os.path.join(OUT, 'ini.bkp')
open(p, 'w', encoding='utf-8', errors='ignore').write(t)
print('写出', p, len(t))


def run(src, tag):
    MSGS = []

    class Sink:
        def OnControlPanelMessage(self, *a):
            s = ' '.join(str(x) for x in a).strip()
            if s and s != 'False':
                MSGS.append(s)

    doc = win32.DispatchEx('Apwn.Document')
    doc.SuppressDialogs = True
    win32.WithEvents(doc, Sink)
    doc.InitFromArchive2(src)
    time.sleep(2)
    doc.Engine.Run2(False)
    t1 = time.time()
    while time.time() - t1 < 600:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.25)
        if time.time() - t1 > 10:
            try:
                if not bool(doc.Engine.IsRunning):
                    break
            except Exception:
                break
    for _ in range(60):
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)

    def g(pp):
        n = doc.Tree.FindNode(pp)
        try:
            return n.Value if n is not None else None
        except Exception:
            return None

    err = [x for x in MSGS if re.search(r'ERROR|Errors|SEVERE|Terminal', x, re.I)]
    print()
    print('===== %s =====' % tag)
    for x in err[:8]:
        print('   ', x[:150])
    IN = ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-132', 'S-210']
    OUTL = ['S-130', 'S-131', 'S-211', 'S-212', 'S-213', 'S-115', 'S-119',
            'S-122', 'S-216', 'S-217', 'S-218', 'S-405C', 'S-406', 'S-209']
    vi = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % x) or 0) for x in IN)
    vo = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % x) or 0) for x in OUTL)
    print('   进=%.2f 出=%.2f 偏差=%.4f%%' % (vi, vo, abs(vi - vo) / vi * 100 if vi else 0))
    for s in ['S-209', 'S-114', 'S-121', 'S-106', 'S-109']:
        print('   %-7s MASS=%-12s T=%s' % (s, g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s),
                                           g(r'\Data\Streams\%s\Output\TEMP_OUT\MIXED' % s)))
    for b in ['T-401', 'T-402', 'T-404', 'T-201']:
        print('   %-7s COND=%-14s REB=%s' % (b, g(r'\Data\Blocks\%s\Output\COND_DUTY' % b),
                                             g(r'\Data\Blocks\%s\Output\REB_DUTY' % b)))
    nn = doc.Tree.FindNode(r'\Data\Streams\S-209\Output\MASSFLOW3')
    if nn is not None:
        items = []
        for i in range(nn.Elements.Count):
            e = nn.Elements.Item(i)
            try:
                v = float(e.Value or 0)
                if abs(v) > 0.01:
                    items.append('%s=%.3f' % (e.Name, v))
            except Exception:
                pass
        print('   S-209 组成:', ', '.join(items))
    if tag == 'BIP':
        doc.SaveAs(r'D:\<化工工作区>\NA-Chemical-10000t_BIP.bkp')
    try:
        doc.Close()
    except Exception:
        pass
    try:
        doc.Quit()
    except Exception:
        pass


run(BASE, 'BASE')
run(p, 'BIP')
print('DONE')
