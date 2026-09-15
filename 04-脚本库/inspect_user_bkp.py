# -*- coding: utf-8 -*-
"""检查用户回存的 bkp：分子结构是否已填、NRTL 参数状态、数据库段"""
import sys, os, re
sys.stdout.reconfigure(encoding='utf-8')

P = r'C:\Users\<用户名>\Desktop\烟酰胺Aspen模型_请补参数后另存.bkp'
print('文件存在:', os.path.exists(P))
if os.path.exists(P):
    print('大小: %d bytes' % os.path.getsize(P))

x = open(P, encoding='utf-8', errors='ignore').read()
print('总长: %d' % len(x))
print()
print('--- 关键字计数 ---')
for k in ['ATOMTYPE', 'ATOMNUM', 'BONDTYPE', 'ATOM1', 'ATOM2',
          'UNIFAC', 'UFGRP', 'GROUPNO', 'MOLEC-STRUCT',
          'BPVAL', 'UVAL1', 'CID1', 'FILE-SYM-NAM', 'DATABANKS',
          'R-PCES', 'AUTO-PARAM', 'LOADDECH']:
    print('  %-14s %d' % (k, x.count(k)))

print()
print('--- NRTL 段原文 ---')
i = x.find('PARAMNAME = NRTL')
print('位置:', i)
if i > 0:
    print(repr(x[i:i + 1600]))

print()
print('--- DATABANKS 段 ---')
j = x.find('DATABANKS')
print(repr(x[max(0, j - 60): j + 400]))

print()
print('--- 分子结构段（前 2 个）---')
for m in list(re.finditer(r'MOLEC-STRUCT"?\s+([A-Za-z0-9_-]+)', x))[:2]:
    s = m.start()
    e = x.find('\n\\', s + 30)
    print('  [%s] @%d: %r' % (m.group(1), s, x[s:min(e if e > 0 else s + 400, s + 400)]))
