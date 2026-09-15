# -*- coding: utf-8 -*-
p = r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\Bulk Chemicals\pfdtut.bkp'
t = open(p, encoding='utf-8', errors='ignore').read()
print('=== BLOCK RSTOIC B2 段 (12463~12909) ===')
print(t[12463:12909])
print()
print('=== 全文搜 REACT / STOIC / RXN ===')
import re
for k in ['REACTIONS', 'REACSET', 'STOIC', 'RXN', 'CONVERSION', 'COEF']:
    print('  %-14s %d' % (k, t.count(k)))
print()
print('=== 含 REACT 的片段 ===')
for m in re.finditer(r'.{0,120}REACT.{0,200}', t):
    print(repr(m.group()))
    print('---')
