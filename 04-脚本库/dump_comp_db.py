# -*- coding: utf-8 -*-
import re

P = r'D:\<化工工作区>\NA_struct.bkp'
t = open(P, encoding='utf-8', errors='ignore').read()
i = t.find('? COMPONENTS MAIN ?')
j = t.find('? ', i + 25)
seg = t[i:j]
seg = re.sub(r'\s+', ' ', seg)
parts = [p.strip() for p in seg.split('/') if p.strip()]
print('组分条数:', len(parts))
print()
print('%-8s %-12s %-26s %-20s %-6s %s' % ('CID', 'ANAME', 'DBNAME1', 'ANAME1', 'TYPE', '其它'))
for p in parts:
    def g(k):
        m = re.search(k + r'\s*=\s*("[^"]*"|\S+)', p)
        return m.group(1).strip('"') if m else '-'
    print('%-8s %-12s %-26s %-20s %-6s' % (g('CID'), g('ANAME'), g('DBNAME1'), g('ANAME1'), g('TYPE')))

# 检查 HENRY 组件的亨利参数是否存在
k = t.find('HENRY-1')
print()
print('--- HENRY-1 段（前 600 字符）---')
print(re.sub(r'[ \t]+', ' ', t[k:k+600]))
