# -*- coding: utf-8 -*-
"""用 Aspen 端口重建拓扑 + 找塔水力学节点"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_最终定稿.bkp'
doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
doc.InitFromArchive2(SRC)
time.sleep(2)
L = []

def g(p):
    n = doc.Tree.FindNode(p)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return None

def names(path, lim=400):
    n = doc.Tree.FindNode(path)
    if n is None:
        return []
    out = []
    try:
        for i in range(min(n.Elements.Count, lim)):
            out.append(n.Elements.Item(i).Name)
    except Exception:
        pass
    return out

# ---- 流股拓扑 ----
sn = doc.Tree.FindNode(r'\Data\Streams')
sids = [sn.Elements.Item(i).Name for i in range(sn.Elements.Count)]
L.append('===== 流股拓扑（Aspen 端口）=====')
topo = {}
for s in sids:
    src = g(r'\Data\Streams\%s\Ports\SOURCE' % s)
    dst = g(r'\Data\Streams\%s\Ports\DEST' % s)
    topo[s] = (src, dst)
    fl = g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s)
    L.append('  %-9s %-34s -> %-34s  %s kg/h' % (
        s, str(src), str(dst), ('%.2f' % float(fl)) if fl is not None else 'None'))

# 反向：块的进出口
L.append('')
L.append('===== 块进出口（按端口反查）=====')
blkn = doc.Tree.FindNode(r'\Data\Blocks')
bids = [blkn.Elements.Item(i).Name for i in range(blkn.Elements.Count)]
bins = {b: [] for b in bids}
bouts = {b: [] for b in bids}
for s, (src, dst) in topo.items():
    if src:
        for b in bids:
            if b in str(src):
                bouts[b].append(s)
    if dst:
        for b in bids:
            if b in str(dst):
                bins[b].append(s)
for b in bids:
    L.append('  %-8s IN : %-46s OUT: %s' % (b, ','.join(bins[b]), ','.join(bouts[b])))

# ---- 塔的水力学/负荷节点名 ----
L.append('')
L.append('===== T-401 Output 全部子节点中的流量类 =====')
alln = names(r'\Data\Blocks\T-401\Output')
L.append('  子节点总数 %d' % len(alln))
for k in alln:
    if re.search(r'VFLOW|LFLOW|MASS|FLOW|DENS|MW|LOAD|HYD|DIAM|AREA|HT|RR|DF|TEMP|PRES', k):
        L.append('    %-16s = %s' % (k, str(g(r'\Data\Blocks\T-401\Output\%s' % k))[:50]))

L.append('')
L.append('===== T-401 Input 全部子节点中的规格类 =====')
allin = names(r'\Data\Blocks\T-401\Input')
L.append('  子节点总数 %d' % len(allin))
for k in allin:
    if re.search(r'NSTAGE|RR|D:F|B:F|FEED|PRES|P_SPEC|FRAC|REFLUX|SPEC|BASIS|CONDENS|REBOIL', k):
        L.append('    %-18s = %s' % (k, str(g(r'\Data\Blocks\T-401\Input\%s' % k))[:50]))

L.append('')
L.append('===== T-201 / T-202 / T-301 / T-302 关键输入 =====')
for b in ['T-201', 'T-202', 'T-301', 'T-302', 'T-401', 'T-402', 'T-403', 'T-404']:
    kv = []
    for k in ['NSTAGE', 'BASIS_RR', 'BASIS_D', 'BASIS_B', 'PRES1', 'CONDENSER', 'REBOILER', 'D:F', 'B:F', 'MODEL', 'SPEC_OPT']:
        v = g(r'\Data\Blocks\%s\Input\%s' % (b, k))
        if v is not None:
            kv.append('%s=%s' % (k, str(v)[:16]))
    L.append('  %-7s %s' % (b, ' '.join(kv)))

open(r'D:\<化工工作区>\_probe\topo_hyd.txt', 'w', encoding='utf-8').write('\n'.join(L))
try:
    doc.Close()
except Exception:
    pass
print('DONE')
