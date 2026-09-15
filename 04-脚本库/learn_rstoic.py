# -*- coding: utf-8 -*-
"""从其它 Aspen 示例学 RSTOIC 多条 STOIC 的写法"""
import re, os, glob

cands = [
    r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\Bulk Chemicals\cumene.bkp',
    r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\Bulk Chemicals\advchex.bkp',
    r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\Bulk Chemicals\pfdtut.bkp',
]
for p in cands:
    if not os.path.exists(p):
        continue
    t = open(p, encoding='utf-8', errors='ignore').read()
    print('=' * 20, os.path.basename(p), flush=True)
    for m in re.finditer(r'\?\s*BLOCK\s+RSTOIC\s+\S+\s*\?', t):
        m2 = re.search(r'\?\s+[A-Z]', t[m.end():])
        end = m.end() + (m2.start() if m2 else 1500)
        seg = t[m.start():end]
        print(re.sub(r'\s+', ' ', seg)[:2000], flush=True)
        print('---', flush=True)
    print(flush=True)
