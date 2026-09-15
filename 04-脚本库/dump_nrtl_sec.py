# -*- coding: utf-8 -*-
import re

P = r'D:\<化工工作区>\NA_struct.bkp'
t = open(P, encoding='utf-8', errors='ignore').read()

for key in ['PARAMETERS BINARY', 'NRTL-1', 'BPVAL', 'BDBANK', 'ESTIMATE']:
    print('%-20s count=%d' % (key, t.count(key)))

i = t.find('PARAMETERS BINARY')
print()
print('--- PARAMETERS BINARY 段（前 1800 字符）---')
print(re.sub(r'[ \t]+', ' ', t[i:i+1800]))
