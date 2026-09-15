# -*- coding: utf-8 -*-
"""找文件里所有 PRES1 / D:F / T-401 出现位置"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')

P = r'D:\<化工工作区>\_probe\p6a.bkp'
t = open(P, encoding='utf-8', errors='ignore').read()
print('文件长', len(t))
print()
print('=== PRES1 全部出现 ===')
for m in re.finditer(r'PRES1', t):
    s = max(0, m.start() - 70)
    print('  @%-8d %s' % (m.start(), repr(t[s:m.start() + 60])))
print()
print('=== D:F 全部出现 ===')
for m in re.finditer(r'D:F', t):
    s = max(0, m.start() - 60)
    print('  @%-8d %s' % (m.start(), repr(t[s:m.start() + 45])))
print()
print('=== T-401 全部出现 ===')
for m in re.finditer(r'T-401', t):
    s = max(0, m.start() - 60)
    print('  @%-8d %s' % (m.start(), repr(t[s:m.start() + 55])))
