# -*- coding: utf-8 -*-
"""找文件里 T-301 / EXTRACT 的所有出现位置"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')

t = open(r'D:\<化工工作区>\_probe\bipA2.bkp', encoding='utf-8', errors='ignore').read()
print('文件长:', len(t))
print()
print('=== "T-301" 全部出现 ===')
for m in re.finditer(r'T-301', t):
    s = max(0, m.start() - 90)
    print('  @%-7d %s' % (m.start(), repr(t[s:m.start() + 90])))
print()
print('=== EXTRACT 全部出现 ===')
for m in re.finditer(r'EXTRACT', t):
    s = max(0, m.start() - 90)
    print('  @%-7d %s' % (m.start(), repr(t[s:m.start() + 70])))
print()
print('=== BLKTYPE 出现统计 ===')
from collections import Counter
c = Counter(re.findall(r'BLKTYPE = "([A-Z0-9]+)"', t))
print(' ', dict(c))
print()
print('=== 是否有其它块清单结构（含 14 的段）===')
for m in re.finditer(r'.{0,60}\b14\b.{0,80}', t[:6000]):
    print('  ', repr(m.group()))
