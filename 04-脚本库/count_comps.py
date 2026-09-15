# -*- coding: utf-8 -*-
import re

for p in [r'D:\<化工工作区>\NA-Chemical-10000t_BIP.bkp', r'D:\<化工工作区>\_probe\ini.bkp']:
    t = open(p, encoding='utf-8', errors='ignore').read()
    i = t.find('? COMPONENTS MAIN ?')
    j = t.find('PROPERTIES PARAMETERS', i)
    seg = t[i:j]
    cids = re.findall(r'CID\s*=\s*("?[^"\s/]+"?)', seg)
    print(p.split('\\')[-1])
    print('   组分条数:', len(cids))
    print('   ', [c.strip('"') for c in cids])
    print('   HENRY 组分:', re.findall(r'HENRY.*?CID\s*=\s*("?[^"\s/]+"?)', t[:t.find('COMPONENTS MAIN') + 3000])[:6])
