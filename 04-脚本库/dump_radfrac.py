# -*- coding: utf-8 -*-
"""dump 全部 RadFrac 段原文（看 PRES1 / D:F 完整写法）"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')

P = r'D:\<化工工作区>\_probe\bipA2.bkp'
t = open(P, encoding='utf-8', errors='ignore').read()
starts = [m.start() for m in re.finditer(r'\?\s*BLOCK\s+', t)]
for i in range(len(starts)):
    s0 = starts[i]
    s1 = starts[i + 1] if i + 1 < len(starts) else len(t)
    m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
    if not m:
        continue
    if m.group(1) != 'RADFRAC':
        continue
    print('=' * 74)
    print(m.group(2), ' 长度', s1 - s0)
    print(t[s0:s1][:1100])
    print()
