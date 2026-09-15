# -*- coding: utf-8 -*-
"""核对 p6a 里 T-401 的 COL-SPECS 实际文本"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')

t = open(r'D:\<化工工作区>\_probe\p6a.bkp', encoding='utf-8', errors='ignore').read()
starts = [s.start() for s in re.finditer(r'\?\s*BLOCK\s+', t)]
for k in range(len(starts) - 1, -1, -1):
    s0 = starts[k]
    s1 = starts[k + 1] if k + 1 < len(starts) else len(t)
    m = re.match(r'\?\s*BLOCK\s+RADFRAC\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
    if m and m.group(1) == 'T-401':
        print(t[s0:s1])
        break
