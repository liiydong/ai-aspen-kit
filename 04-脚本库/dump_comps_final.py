# -*- coding: utf-8 -*-
"""核对模型里每个组分到底是什么（CID / DBNAME1 / ANAME1 / CAS）"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')

P = r'D:\<化工工作区>\NA-Chemical-10000t_最终版.bkp'
t = open(P, encoding='utf-8', errors='ignore').read()
i = t.find('? COMPONENTS MAIN ?')
j = t.find('? COMPONENTS "COMP-LIST"', i)
seg = t[i:j]
# 按 '/' 切记录
raw = re.sub(r'\s+', ' ', seg)
parts = re.split(r'\s/\s', raw)
print('记录数:', len(parts))
print()
print('%-10s %-12s %-24s %-16s %-10s' % ('CID', 'OUTNAME', 'DBNAME1', 'ANAME1', 'TYPE'))
for e in parts:
    cid = re.search(r'CID = "?([A-Za-z0-9-]+)"?', e)
    out = re.search(r'OUTNAME = "?([A-Za-z0-9-]+)"?', e)
    db = re.search(r'DBNAME1 = "([^"]+)"', e)
    a1 = re.search(r'ANAME1 = "([^"]+)"', e)
    tp = re.search(r'TYPE = (\w+)', e)
    if cid:
        print('%-10s %-12s %-24s %-16s %-10s' % (
            cid.group(1),
            out.group(1) if out else '-',
            db.group(1) if db else '-',
            a1.group(1) if a1 else '-',
            tp.group(1) if tp else '-'))

# 找 4-CP / 3-CP 的纯组分数据段
print()
print('=== 与 4-CP / 3-CP 有关的 DSET 记录 ===')
for m in list(re.finditer(r'DSET[^\n]{0,120}(4-CP|3-CP)[^\n]{0,80}', t))[:12]:
    print('  ', repr(m.group()[:160]))
