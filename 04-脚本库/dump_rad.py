# -*- coding: utf-8 -*-
import re

p = r'D:\<化工工作区>\NA-Chemical-10000t_BIP.bkp'
t = open(p, encoding='utf-8', errors='ignore').read()

for bid in ['T-401', 'T-402', 'T-404']:
    m = re.search(r'\?\s*BLOCK\s+RADFRAC\s+"?%s"?\s*\?' % re.escape(bid), t)
    m2 = re.search(r'\?\s*BLOCK\s', t[m.end():])
    end = m.end() + m2.start() if m2 else m.end() + 5000
    print('=' * 20, bid)
    print(t[m.start():end][:1800])
    print()
