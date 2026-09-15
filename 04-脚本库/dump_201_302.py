# -*- coding: utf-8 -*-
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
t = open(r'D:\<化工工作区>\NA-Chemical-10000t_s3.bkp', encoding='utf-8', errors='ignore').read()
starts = [s.start() for s in re.finditer(r'\?\s*BLOCK\s+', t)]
out = []
for k in range(len(starts)):
    s0 = starts[k]
    s1 = starts[k + 1] if k + 1 < len(starts) else len(t)
    m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
    if m and m.group(2) in ('T-201', 'T-302'):
        out.append('=' * 30 + ' ' + m.group(2))
        seg = t[s0:s1]
        out.append('段长 %d' % len(seg))
        out.append(seg)
        out.append('')
open(r'D:\<化工工作区>\_probe\t201_t302.txt', 'w', encoding='utf-8').write('\n'.join(out))
print('DONE')
