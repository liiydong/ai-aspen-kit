# -*- coding: utf-8 -*-
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
t = open(r'D:\<化工工作区>\NA-Chemical-10000t_s6.bkp', encoding='utf-8', errors='ignore').read()
L = []
i = t.find('; \n')
L.append('=== 注册表头部 ===')
L.append(t[max(0, i - 120):i + 220])
L.append('')
L.append('=== 注册表条目 ===')
for m in re.finditer(r'>VERSION 0\n([^\n]+)\n([^\n]+)\n([^\n]+)\n([^\n]+)\n', t):
    if m.group(1) in ('E-102', 'E-105', 'E-601', 'P-301', 'M-502', 'C-101', 'M-101'):
        L.append('  >> %s | %s | %s | %s' % (m.group(1), m.group(2), m.group(3), m.group(4)))
starts = [s.start() for s in re.finditer(r'\?\s*BLOCK\s+', t)]
for k in range(len(starts)):
    s0 = starts[k]
    s1 = min(starts[k + 1] if k + 1 < len(starts) else s0 + 900, s0 + 900)
    m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
    if m and m.group(2) in ('E-102', 'E-105', 'P-301', 'C-101'):
        L.append('=' * 20 + ' ' + m.group(2))
        L.append(t[s0:s1])
open(r'D:\<化工工作区>\_probe\heater_fmt2.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('DONE')
