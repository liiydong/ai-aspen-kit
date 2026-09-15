# -*- coding: utf-8 -*-
"""估算测试：UNIF-DMD 方法"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')

BASE = r'D:\<化工工作区>'
SRC = os.path.join(BASE, 'NA_struct.bkp')
x = open(SRC, encoding='utf-8', errors='ignore').read()

# 设估算方法为 UNIF-DMD（Dortmund 修正版，基团表更全）
n1 = 0
for old, new in [
    ('ESTIMATE = NO', 'ESTIMATE = YES'),
    ('ALLONLY = "NONE"', 'ALLONLY = "ALL"'),
]:
    if old in x:
        x = x.replace(old, new)
        n1 += 1
print('基础开关替换:', n1)

# 在 Estimation\Estimate\Input 段里插入/修改 METHODG
i = x.find('? PROPERTIES ESTIMATION')
print('ESTIMATION 段位置:', i)
if i > 0:
    print(repr(x[i:i + 700]))

OUT = os.path.join(BASE, 'NA_est_dmd.bkp')
open(OUT, 'w', encoding='utf-8', errors='ignore').write(x)
print('写出', os.path.basename(OUT), os.path.getsize(OUT))
