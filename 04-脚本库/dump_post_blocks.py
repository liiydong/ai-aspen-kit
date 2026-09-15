# -*- coding: utf-8 -*-
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
t = open(r'D:\<化工工作区>\NA-Chemical-10000t_s5.bkp', encoding='utf-8', errors='ignore').read()
starts = [s.start() for s in re.finditer(r'\?\s*BLOCK\s+', t)]
targets = ('E-601', 'E-602', 'E-603', 'C-601', 'D-601', 'M-501', 'M-502', 'M-601')
L = []
for k in range(len(starts)):
    s0 = starts[k]
    s1 = starts[k + 1] if k + 1 < len(starts) else len(t)
    m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
    if m and m.group(2) in targets:
        L.append('=' * 22 + ' ' + m.group(2) + ' (len=%d)' % (s1 - s0))
        L.append(t[s0:s1])
        L.append('')
open(r'D:\<化工工作区>\_probe\post_blocks.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('DONE')
