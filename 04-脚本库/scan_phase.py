# -*- coding: utf-8 -*-
"""在 bkp 文本中搜索相稳定性/亨利/相态相关字段"""
import re

P = r'D:\<化工工作区>\NA-Chemical-10000t_BIP2.bkp'
t = open(P, encoding='utf-8', errors='ignore').read()

KEYS = ['STAB', 'FREE-WATER', 'FREEWATER', 'NPHASE', 'VLCHECK', 'CHKPHASE',
        'HENRY', 'PHASECHK', 'VALIDPHASE', 'VLL', 'L1-COMPS', 'L2-COMPS',
        'PHASE-STAB', 'STABILITY']
print('%-14s %s' % ('关键字', '出现次数'))
for k in KEYS:
    c = len(re.findall(re.escape(k), t, re.I))
    print('%-14s %d' % (k, c))

print()
print('=== HENRY 相关上下文 ===')
for m in re.finditer(r'.{0,80}HENRY.{0,160}', t):
    print('  ', re.sub(r'\s+', ' ', m.group(0))[:230])
    print()

print('=== T-301 SEP 段 ===')
m = re.search(r'\?\s*BLOCK\s+SEP\s+"?T-301"?\s*\?', t)
if m:
    m2 = re.search(r'\?\s*BLOCK\s', t[m.end():])
    end = m.end() + (m2.start() if m2 else 3000)
    print(t[m.start():end][:2200])
else:
    print('  未找到 SEP T-301')
