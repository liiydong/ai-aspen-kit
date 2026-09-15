# -*- coding: utf-8 -*-
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
t = open(r'D:\<化工工作区>\NA-Chemical-10000t_s8_s.bkp', encoding='utf-8', errors='ignore').read()
L = []
for sid in ['S-101', 'S-103', 'S-111']:
    L.append('===== %s 出现位置 =====' % sid)
    for m in re.finditer(re.escape(sid), t):
        i = m.start()
        seg = re.sub(r'\s+', ' ', t[max(0, i - 120):i + 160])
        L.append('  @%d  ...%s...' % (i, seg))
        if len(L) > 60:
            break
    L.append('')
open(r'D:\<化工工作区>\_probe\find_ids.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('DONE')
