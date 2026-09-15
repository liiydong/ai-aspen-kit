# -*- coding: utf-8 -*-
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
P = r'D:\<化工工作区>\NA-Chemical-10000t_Submit.bkp'
t = open(P, encoding='utf-8', errors='ignore').read()
flat = re.sub(r'\s+', ' ', t)
pairs = re.findall(r'PARAMNAME2\s*=\s*NRTL\s+CID1\s*=\s*"?([A-Za-z0-9-]+)"?\s+CID2\s*=\s*"?([A-Za-z0-9-]+)"?', flat)
L = []
L.append('===== 文件中 NRTL 二元对：%d 条 =====' % len(pairs))
for a, b in pairs:
    L.append('   %-8s - %-8s' % (a, b))
# 各自的来源
for m in re.finditer(r'BPVAL PARAMNAME2 = NRTL CID1 = "?([A-Za-z0-9-]+)"? CID2 = "?([A-Za-z0-9-]+)"?(.{0,400})', flat):
    src = re.findall(r'VAL\d+ = "([^"]+)"', m.group(3))
    L.append('   %-8s - %-8s  来源: %s' % (m.group(1), m.group(2), sorted(set(src))))
# 所有物性方法的 BDBANK 行
L.append('')
L.append('===== BDBANK 出现处 =====')
for m in re.finditer(r'BDBANK\s*=\s*\((.{0,220}?)\)\s*NEL', flat):
    L.append('   (%s)' % m.group(1))
# DATABANKS 段原文
m = re.search(r'\?\s*DATABANKS\s*\?(.{0,600})', flat)
L.append('')
L.append('===== DATABANKS 段原文 =====')
L.append(m.group(1) if m else 'not found')
# ODATABANKS
m2 = re.search(r'\?\s*ODATABANKS\s*\?(.{0,300})', flat)
L.append('')
L.append('===== ODATABANKS 段 =====')
L.append(m2.group(1) if m2 else 'not found')
open(r'D:\<化工工作区>\_probe\bip_pairs.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('\n'.join(L))
