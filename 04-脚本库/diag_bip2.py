# -*- coding: utf-8 -*-
"""诊断：NRTL 二元参数到底检索到几对、缺哪些、UNIFAC 基团是否可用"""
import sys, re, os
sys.stdout.reconfigure(encoding='utf-8')
P = r'D:\<化工工作区>\NA-Chemical-10000t_Submit.bkp'
t = open(P, encoding='utf-8', errors='ignore').read()
L = []
L.append('文件: %s  %d 字节' % (os.path.basename(P), len(t)))

# 1) DATABANKS / BDBANK
for key in ['DATABANKS', 'BDBANK', 'FILE-SYM-NAM', 'IN-UNITS', 'RUN-CLASS']:
    L.append('')
    L.append('===== %s =====' % key)
    for m in re.finditer(re.escape(key), t):
        i = m.start()
        L.append('  pos %7d : %s' % (i, re.sub(r'\s+', ' ', t[max(0, i - 60):i + 320])))

# 2) 所有 PARAMETERS BINARY 段
L.append('')
L.append('===== PARAMETERS BINARY 段 =====')
for m in re.finditer(r'PARAMETERS\s+BINARY', t):
    i = m.start()
    j = t.find('\n? ', i)
    j = j if j > 0 else i + 4000
    seg = t[i:j]
    L.append('---- pos %d  len %d' % (i, len(seg)))
    L.append(re.sub(r'[ \t]+', ' ', seg[:3000]))
    L.append('')

open(r'D:\<化工工作区>\_probe\bip_diag.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('\n'.join(L[:200]))
