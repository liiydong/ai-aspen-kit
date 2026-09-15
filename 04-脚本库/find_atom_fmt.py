# -*- coding: utf-8 -*-
"""在示例 bkp 里搜 原子/键/UNIFAC 基团 的原始格式"""
import sys, re, os
sys.stdout.reconfigure(encoding='utf-8')

FILES = [
    (r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\Bulk Chemicals\Distillation\3phase.bkp', '3phase'),
    (r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\Batch Modeling\Batch Distillation\WaterMethanol.bkp', 'WaterMethanol'),
    (r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\Bulk Chemicals\pfdtut.bkp', 'pfdtut'),
    (r'D:\<化工工作区>\NA_work.bkp', 'OURS'),
]

KEYS = ['ATOMTYPE', 'BONDTYPE', 'UNIFAC', 'UFGRP', 'MOLEC-STRUCT', 'GROUPNO']

for p, tag in FILES:
    print('=' * 74)
    if not os.path.exists(p):
        print('缺失', tag); continue
    x = open(p, encoding='utf-8', errors='ignore').read()
    print(tag, '|', len(x), 'bytes')
    for k in KEYS:
        n = x.count(k)
        print('   %-14s 出现 %d 次' % (k, n))
    # 若有 ATOM 数据，dump 片段
    for k in ['ATOMTYPE', 'UNIFAC', 'UFGRP']:
        i = x.find(k)
        if i > 0:
            print('   --- %s 片段 ---' % k)
            print('   ', repr(x[max(0, i - 200): i + 500]))
    print()
