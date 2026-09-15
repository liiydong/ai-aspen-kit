# -*- coding: utf-8 -*-
"""检查 v14 改动：FLOWSHEET 中 T-301 写法、各 BLOCK 段头、D:F 出现处"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')

for tag, P in [('源 bipA2', r'D:\<化工工作区>\_probe\bipA2.bkp'),
               ('v14', r'D:\<化工工作区>\NA-Chemical-10000t_v14.bkp')]:
    t = open(P, encoding='utf-8', errors='ignore').read()
    print('=' * 74)
    print(tag, len(t))
    print('--- FLOWSHEET 中 T-301 上下文 ---')
    i = t.find('"T-301"')
    print(repr(t[i-160:i+120]))
    print('--- 所有 BLOCK 段头 ---')
    for m in re.finditer(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', t):
        print('   %-10s %s' % (m.group(1), m.group(2)), end='')
    print()
    print('--- D:F 出现处 ---')
    for m in re.finditer(r'.{0,30}D:F = [\d.]+.{0,20}', t):
        print('   ', repr(m.group()))
    print()
