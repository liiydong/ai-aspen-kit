# -*- coding: utf-8 -*-
import re

P = r'D:\<化工工作区>\NA-Chemical-10000t_BIP2.bkp'
t = open(P, encoding='utf-8', errors='ignore').read()

def ctx(key, before=90, after=170, limit=8):
    print('=== %s ===' % key)
    seen = set()
    n = 0
    for m in re.finditer(re.escape(key), t):
        s = re.sub(r'\s+', ' ', t[max(0, m.start()-before): m.start()+after])
        if s in seen:
            continue
        seen.add(s)
        print('  ...', s)
        n += 1
        if n >= limit:
            break
    print()

for k in ['FREE-WATER', 'VLL', 'STABILITY', 'STAB']:
    ctx(k)
