# -*- coding: utf-8 -*-
"""dump q3_s.bkp 里已有块段落的完整原文（作为模板）"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
P = r'D:\<化工工作区>\NA-Chemical-10000t_q3_s.bkp'
t = open(P, encoding='utf-8', errors='ignore').read()
starts = [m.start() for m in re.finditer(r'\?\s*BLOCK\s+', t)]
for k in range(len(starts)):
    s0 = starts[k]
    s1 = starts[k + 1] if k + 1 < len(starts) else min(len(t), s0 + 900)
    print('=' * 76)
    print(t[s0:s1][:900])
    print()
