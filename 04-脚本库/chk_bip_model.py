# -*- coding: utf-8 -*-
import re, os

p = r'D:\<化工工作区>\NA-Chemical-10000t_BIP.bkp'
t = open(p, encoding='utf-8', errors='ignore').read()
print('size', len(t))
print('BLOCK 段:', len(re.findall(r'\?\s*BLOCK\s', t)))
print('RUN-CLASS:', re.findall(r'RUN-CLASS = (\w+)', t))
i = t.find('PARAMNAME = NRTL')
j = t.find('PARAMNAME =', i + 20)
seg = t[i:j] if j > i else t[i:i+40000]
print('UVAL =', seg.count('UVAL'), ' BPVAL =', seg.count('BPVAL'))
flat = re.sub(r'\s+', ' ', seg)
print('参数对:')
for m in re.finditer(r'CID1 = ("?[^"\s]+"?)\s+CID2 = ("?[^"\s]+"?)', flat):
    print('   ', m.group(1), '/', m.group(2))
print('组分条数:', len(re.findall(r'CID = ', t[t.find('COMPONENTS MAIN'):t.find('COMPONENTS MAIN') + 3000])))
print('ESTIMATE:', re.findall(r'ESTIMATE = (\w+)', t)[:4])
