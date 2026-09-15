# -*- coding: utf-8 -*-
"""探测 xlsx 报告的列切分宽度"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
P = r'D:\<化工工作区>\_probe\report_cols.tsv'
rows = {}
for ln in open(P, encoding='utf-8').read().split('\n'):
    if not ln.strip():
        continue
    p = ln.split('\t')
    rows[int(p[0])] = p[1:]

out = []
for r in [4140, 4143, 4150, 4166, 4170, 4181, 4797]:
    if r in rows:
        c = rows[r]
        out.append('R%d  ncol=%d' % (r, len(c)))
        for i, v in enumerate(c):
            out.append('   [%d] len=%d  %r' % (i, len(v), v))

# 统计 B 列长度分布
from collections import Counter
cnt = Counter()
for r, c in rows.items():
    if len(c) >= 2:
        cnt[len(c[1])] += 1
out.append('')
out.append('B-len distribution: ' + str(sorted(cnt.items())[:40]))
print('\n'.join(out))
