# -*- coding: utf-8 -*-
"""① 核查 NA_work 的 NRTL 记录原文 ② PENG-ROB 对照运行"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')

BASE = r'D:\<化工工作区>'
SRC = os.path.join(BASE, 'NA_work.bkp')
x = open(SRC, encoding='utf-8', errors='ignore').read()

print('===== ① NRTL 段原文 =====')
i = x.find('PARAMNAME = NRTL')
print('PARAMNAME = NRTL 位置:', i)
if i > 0:
    seg = x[i:i + 3000]
    j = seg.find('\n\\ ? ')
    if j > 0:
        seg = seg[:j]
    print(repr(seg[:2200]))
print()
print('含 BPVAL:', x.count('BPVAL'))
print('含 "CID1 = H2O":', x.count('CID1 = H2O'))
print('含 H2O CID2:', x.count('H2O CID2'))
print('含 LLE-ASPEN:', x.count('LLE-ASPEN'))
print('含 NRTL-1:', x.count('NRTL-1'))
