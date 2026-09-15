# -*- coding: utf-8 -*-
"""挖 DATABANKS / BINARY / MOLEC-STRUCT 三段原文"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')

P = r'D:\<化工工作区>\NA-Chemical-10000t_v13_s.bkp'
t = open(P, encoding='utf-8', errors='ignore').read()

def seg(name):
    i = t.find(name)
    if i < 0:
        return None
    return i

# DATABANKS 原始
print('=' * 70)
i = seg('? DATABANKS ?')
print('DATABANKS @', i)
print(repr(t[i:i+220]))

# BINARY / PARAMETERS 区
print('=' * 70)
i = seg('PARAMETERS BINARY')
print('PARAMETERS BINARY @', i)
print(repr(t[i-200:i+1600]))

# MOLEC-STRUCT 3-CP
print('=' * 70)
i = seg('"MOLEC-STRUCT" "3-CP"')
print('3-CP struct @', i)
print(repr(t[i:i+700]))

# MOLEC-STRUCT NAM
print('=' * 70)
i = seg('"MOLEC-STRUCT" NAM')
print('NAM struct @', i)
print(repr(t[i:i+700]))

# 统计 DATABANKS 后紧跟什么
print('=' * 70)
print('UNIFAC-GROUP 段:')
i = seg('"UNIFAC-GROUP"')
print(repr(t[i-60:i+600]) if i >= 0 else '未找到')
