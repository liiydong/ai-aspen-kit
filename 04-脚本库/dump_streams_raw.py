# -*- coding: utf-8 -*-
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
t = open(r'D:\<化工工作区>\NA-Chemical-10000t_s8_s.bkp', encoding='utf-8', errors='ignore').read()
L = []
for sid in ['S-101', 'S-102', 'S-103', 'S-111', 'S-124', 'S-210']:
    i = t.find('? STREAM MATERIAL "%s"' % sid)
    L.append('=== %s  i=%d' % (sid, i))
    if i < 0:
        continue
    j = t.find('? STREAM MATERIAL', i + 5)
    k = t.find('? BLOCK', i + 5)
    cands = [q for q in (j, k) if q > 0]
    e = min(cands) if cands else i + 600
    L.append(t[i:min(e, i + 700)])
    L.append('')
open(r'D:\<化工工作区>\_probe\streams_raw.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('DONE')
