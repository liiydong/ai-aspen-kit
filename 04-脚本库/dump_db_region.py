# -*- coding: utf-8 -*-
import re

P = r'D:\<化工工作区>\NA_struct.bkp'
t = open(P, encoding='utf-8', errors='ignore').read()
print('size', len(t))

# DATABANKS / ODATABANKS 区域
for key in ['DATABANKS', 'ODATABANKS', 'FILE-SYM-NAM', 'AUTO-PARAM']:
    print('%-16s count=%d' % (key, t.count(key)))

i = t.find('DATABANKS')
print()
print('--- DATABANKS 首次出现处（前后 500 字符）---')
print(repr(t[max(0, i-300): i+500]))
