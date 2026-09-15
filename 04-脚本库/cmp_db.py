# -*- coding: utf-8 -*-
"""对比参考模型与我们的 DATABANKS / NRTL 二元参数段"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')

REF = r'D:\<化工工作区>\_probe\ref_multi.bkp'
USR = r'D:\<化工工作区>\NA-Chemical-10000t_v13_s.bkp'

for tag, p in [('参考', REF), ('我们', USR)]:
    t = open(p, encoding='utf-8', errors='ignore').read()
    print('=' * 72)
    print('%s  大小=%d' % (tag, len(t)))
    i = t.find('? DATABANKS ?')
    print('--- DATABANKS 段原始 ---')
    print(repr(t[i:i+330]))
    print()

# 参考模型：NRTL-1 段有没有数据记录
t = open(REF, encoding='utf-8', errors='ignore').read()
i = t.find('"NRTL-1"')
print('=' * 72)
print('参考 NRTL-1 @', i)
if i >= 0:
    print(repr(t[i:i+1200]))
else:
    print('  参考模型没有 NRTL-1 段')

# 参考模型用的什么物性方法
print('=' * 72)
m = re.search(r'GBASEOPSET\s*=\s*(\S+)', t)
print('参考 GBASEOPSET =', m.group(1) if m else '?')
for m in re.finditer(r'PARAM\s+BASE\s*=\s*(\S+)', t):
    print('  参考 PARAM BASE =', m.group(1))
