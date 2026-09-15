# -*- coding: utf-8 -*-
"""dump T-301(Extract) 与 T-401 段原文"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')

P = r'D:\<化工工作区>\_probe\bipA2.bkp'
t = open(P, encoding='utf-8', errors='ignore').read()
starts = [m.start() for m in re.finditer(r'\?\s*BLOCK\s+', t)]
for i in range(len(starts) - 1, -1, -1):
    s0 = starts[i]
    s1 = starts[i + 1] if i + 1 < len(starts) else len(t)
    m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
    if not m:
        continue
    bid = m.group(2)
    if bid in ('T-301', 'T-401', 'T-201'):
        print('=' * 74)
        print(bid, '  段长', s1 - s0)
        print(t[s0:s1][:1300])
        print()
