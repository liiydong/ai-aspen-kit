# -*- coding: utf-8 -*-
"""解析 FLOWSHEET 拓扑 + dump 塔水力学节点 + 查 S-207"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_最终定稿.bkp'
t = open(SRC, encoding='utf-8', errors='ignore').read()
L = []

# ---- a) 解析 FLOWSHEET ----
m = re.search(r'\?\s*FLOWSHEET\s+GLOBAL\s*\?', t)
s0 = m.end()
m2 = re.search(r'\n\?\s*[A-Z]', t[s0:])
s1 = s0 + m2.start() if m2 else len(t)
fs = t[s0:s1]
flat = re.sub(r'\s+', ' ', fs)
recs = re.split(r'(?=BLOCK\s+BLKID\s*=)', flat)
L.append('===== FLOWSHEET 拓扑（%d 条）=====' % (len(recs) - 1))
conns = {}
for r in recs:
    m3 = re.match(r'BLOCK\s+BLKID\s*=\s*"([^"]+)"\s+BLKTYPE\s*=\s*"([^"]+)"(?:\s+MDLTYPE\s*=\s*"([^"]*)")?', r)
    if not m3:
        continue
    bid, bt = m3.group(1), m3.group(2)
    ins = re.search(r'IN\s*=\s*\((.*?)\)', r)
    outs = re.search(r'OUT\s*=\s*\((.*?)\)', r)
def streams_of(seg):
    if not seg:
        return []
    return re.findall(r'"([^"]+)"', seg)
order = []
for r in recs:
    m3 = re.match(r'BLOCK\s+BLKID\s*=\s*"([^"]+)"\s+BLKTYPE\s*=\s*"([^"]+)"', r)
    if not m3:
        continue
    bid, bt = m3.group(1), m3.group(2)
    ins = streams_of((re.search(r'IN\s*=\s*\((.*?)\)', r) or [None, ''])[1])
    outs = streams_of((re.search(r'OUT\s*=\s*\((.*?)\)', r) or [None, ''])[1])
    conns[bid] = (bt, ins, outs)
    order.append(bid)
    L.append('  %-8s %-9s IN=[%s]  OUT=[%s]' % (bid, bt, ','.join(ins), ','.join(outs)))
L.append('  共 %d 个块' % len(conns))

# 每条流股的来源/去向
srcmap, dstmap = {}, {}
for bid, (bt, ins, outs) in conns.items():
    for s in ins:
        dstmap.setdefault(s, []).append(bid)
    for s in outs:
        srcmap.setdefault(s, []).append(bid)
alls = sorted(set(list(srcmap) + list(dstmap)))
orphan = [s for s in alls if not srcmap.get(s) and not dstmap.get(s)]
unconn_in = [s for s in alls if not srcmap.get(s)]
L.append('')
L.append('===== 连接完整性 =====')
L.append('  流股总数 %d' % len(alls))
L.append('  无来源(应=外部进料): %s' % ', '.join(unconn_in))
L.append('  无去向: %s' % ', '.join(s for s in alls if not dstmap.get(s)))
L.append('  孤立流股: %s' % (', '.join(orphan) if orphan else '无'))
# 多条去向
L.append('  去向多于一处的流股: %s' % ', '.join('%s->%s' % (s, ','.join(v)) for s, v in dstmap.items() if len(v) > 1))

# ---- S-207 ----
L.append('')
L.append('===== S-207 在文本中的位置 =====')
idx = [m.start() for m in re.finditer(r'S-207', t)]
L.append('  出现 %d 次: %s' % (len(idx), idx[:10]))
for i in idx[:4]:
    L.append('   %r' % re.sub(r'\s+', ' ', t[max(0, i - 120):i + 160]))

doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
doc.InitFromArchive2(SRC)
time.sleep(2)

def g(p):
    n = doc.Tree.FindNode(p)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return None

def names(path):
    n = doc.Tree.FindNode(path)
    if n is None:
        return []
    out = []
    try:
        for i in range(n.Elements.Count):
            out.append(n.Elements.Item(i).Name)
    except Exception:
        pass
    return out

# ---- b) 水力学与温度节点 ----
for b in ['T-201', 'T-302', 'T-401', 'T-402', 'T-404']:
    L.append('')
    L.append('===== %s Output 中 VFLOW/LFLOW/DENS/TEMP/AREA/DIAM/RR 类 =====' % b)
    alln = names(r'\Data\Blocks\%s\Output' % b)
    L.append('  子节点总数 %d' % len(alln))
    for k in alln:
        if re.search(r'VFLOW|LFLOW|DENS|MW|_TEMP|TEMPS|DIAM|AREA|REFLUX|_RR|DUTY|FLOW', k):
            v = g(r'\Data\Blocks\%s\Output\%s' % (b, k))
            if v is not None and str(v) not in ('0', '0.0', '', 'None'):
                L.append('    %-16s = %s' % (k, str(v)[:52]))

# 模型类型
L.append('')
L.append('===== 块模型类型（文本 BLKTYPE）=====')
for bid in sorted(conns):
    L.append('  %-8s %s' % (bid, conns[bid][0]))

open(r'D:\<化工工作区>\_probe\topo2.txt', 'w', encoding='utf-8').write('\n'.join(L))
try:
    doc.Close()
except Exception:
    pass
print('DONE')
