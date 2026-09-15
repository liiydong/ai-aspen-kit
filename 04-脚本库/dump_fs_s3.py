# -*- coding: utf-8 -*-
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
t = open(r'D:\<化工工作区>\NA-Chemical-10000t_s3.bkp', encoding='utf-8', errors='ignore').read()
out = []

# FLOWSHEET 段
i = t.find('"DEF-STREAM"')
j = t.find('? ', i + 20)
fs = re.sub(r'\s+', ' ', t[i:j])
out.append('=== FLOWSHEET 段 ===')
for m in re.finditer(r'BLOCK BLKID = "([^"]+)"[^\\]*?IN = \( ([^)]*) \) OUT = \( ([^)]*) \)', fs):
    out.append('  %-7s IN : %-40s OUT: %s' % (m.group(1), m.group(2), m.group(3)))
out.append('')

# 全部 STREAM 段
out.append('=== STREAM 段清单 ===')
names = []
for m in re.finditer(r'\?\s*STREAM\s+MATERIAL\s+"([^"]+)"', t):
    names.append(m.group(1))
out.append('  共 %d 条: %s' % (len(names), ' '.join(sorted(names))))
out.append('')

# 新增流股段原文
targets = ['S-401', 'S-402', 'S-210', 'S-211', 'S-212', 'S-213', 'S-214', 'S-201', 'S-202', 'S-203']
out.append('=== 新增流股段原文 ===')
starts = [s.start() for s in re.finditer(r'\?\s*STREAM\s+', t)]
for k in range(len(starts)):
    s0 = starts[k]
    s1 = starts[k + 1] if k + 1 < len(starts) else t.find('? BLOCK', s0)
    if s1 < 0:
        s1 = len(t)
    m = re.match(r'\?\s*STREAM\s+\w+\s+"?([A-Za-z0-9-]+)"?', t[s0:s1])
    if m and m.group(1) in targets:
        out.append('--- %s ---' % m.group(1))
        out.append(t[s0:s1])
open(r'D:\<化工工作区>\_probe\fs_s3.txt', 'w', encoding='utf-8').write('\n'.join(out))
print('DONE')
