# -*- coding: utf-8 -*-
"""dump r1 文件的行，定位报错段"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
P = r'D:\<化工工作区>\NA-Chemical-10000t_r1.bkp'
lines = open(P, encoding='utf-8', errors='ignore').read().split('\n')
print('总行数', len(lines))
for a, b in [(140, 200), (238, 252)]:
    print('=' * 70)
    for i in range(a, min(b, len(lines))):
        print('%4d| %s' % (i + 1, lines[i][:110]))
