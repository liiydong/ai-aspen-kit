# -*- coding: utf-8 -*-
"""列出 q3_s.bkp 里主要段落的顺序与偏移，确定新块段落该插在哪"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
P = r'D:\<化工工作区>\NA-Chemical-10000t_q3_s.bkp'
t = open(P, encoding='utf-8', errors='ignore').read()
print('文件长', len(t))
print()
pat = re.compile(r'\?\s*(FLOWSHEET|PROPERTIES|COMPONENTS|SOLVE|STREAM\s+MATERIAL|BLOCK\s+[A-Z0-9]+|SETUP|ODATABANKS|DATABANKS|CURRENCY|"STREAM-PRICE"|REACTIONS)\b')
seen = []
for m in pat.finditer(t):
    key = m.group(1)
    seen.append((m.start(), key))
# 只打印关键位置的分布
print('前 60 个段标记:')
for off, k in seen[:60]:
    print('  @%-8d %s' % (off, k))
print('...')
print('共', len(seen), '个标记')
print()
# 关键位置
for name in ['? STREAM', '? BLOCK', '? PROPERTIES MAIN', '? PROPERTIES "OPTION-SETS"']:
    idxs = [mm.start() for mm in re.finditer(re.escape(name), t)]
    print('%-28s 出现 %d 次，位置 %s' % (name, len(idxs), idxs[:6]))
