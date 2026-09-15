# -*- coding: utf-8 -*-
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
t = open(r'D:\<化工工作区>\NA-Chemical-10000t_s3.bkp', encoding='utf-8', errors='ignore').read()
out = []

i = t.find('? FLOWSHEET GLOBAL ?')
print('FLOWSHEET 起始', i)
j = t.find('? ', i + 25)
seg = t[i:j]
fs = re.sub(r'\s+', ' ', seg)
out.append('=== FLOWSHEET 段 (len=%d) ===' % len(seg))
for m in re.finditer(r'BLOCK BLKID = "([^"]+)" BLKTYPE = "([^"]+)" MDLTYPE = "([^"]+)" IN = \( ([^)]*) \) OUT = \( ([^)]*) \)', fs):
    out.append('  %-8s %-9s IN : %-34s OUT: %s' % (m.group(1), m.group(2), m.group(4), m.group(5)))
out.append('')
out.append('匹配到 %d 条块' % len(re.findall(r'BLOCK BLKID', fs)))
out.append('')
out.append('=== FLOWSHEET 原文（前 2600 字符）===')
out.append(fs[:2600])
open(r'D:\<化工工作区>\_probe\fs_s3.txt', 'w', encoding='utf-8').write('\n'.join(out))
print('DONE')
