# -*- coding: utf-8 -*-
"""精确 dump：3phase 的 DATABANKS 上下文 + WaterMethanol 的 MOLEC-STRUCT / UNIFAC-GROUP"""
import sys, re, os
sys.stdout.reconfigure(encoding='utf-8')

def dump(path, tag):
    print('=' * 74)
    print(tag)
    print('=' * 74)
    x = open(path, encoding='utf-8', errors='ignore').read()

    # DATABANKS 上下文
    i = x.find('DATABANKS')
    if i >= 0:
        print('--- DATABANKS 前 250 / 后 700 字符 ---')
        print(repr(x[max(0, i - 250): i + 700]))
    print()

    # MOLEC-STRUCT 段
    for m in re.finditer(r'MOLEC-STRUCT"?\s+([A-Za-z0-9_-]+)', x):
        s = m.start()
        e = x.find('? PROPERTIES', s + 20)
        if e < 0:
            e = s + 700
        print('--- MOLEC-STRUCT [%s] @%d ---' % (m.group(1), s))
        print(repr(x[s: min(e, s + 700)]))
        print()
        break

    # UNIFAC-GROUP 段
    k = x.find('UNIFAC-GROUP')
    print('--- UNIFAC-GROUP ---')
    if k >= 0:
        print(repr(x[k: k + 600]))
    else:
        print('无')
    print()


dump(r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\Bulk Chemicals\Distillation\3phase.bkp', '3phase (Bulk Chemicals)')
dump(r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\Batch Modeling\Batch Distillation\WaterMethanol.bkp', 'WaterMethanol')

# 我们的文件对照
print('=' * 74)
print('我们的文件 NA_work.bkp')
print('=' * 74)
x = open(r'D:\<化工工作区>\NA_work.bkp', encoding='utf-8', errors='ignore').read()
print('含 MOLEC-STRUCT:', 'MOLEC-STRUCT' in x)
print('含 UNIFAC-GROUP:', 'UNIFAC-GROUP' in x)
i = x.find('DATABANKS')
print('DATABANKS 上下文:', repr(x[max(0, i - 200): i + 250]))
