# -*- coding: utf-8 -*-
import re

p = r'D:\<化工工作区>\NA-Chemical-10000t_工作版.bkp'
t = open(p, encoding='utf-8', errors='ignore').read()

print('=== 全部 BLOCK 段 ===')
secs = [(m.start(), re.sub(r'\s+', ' ', m.group().strip())) for m in
        re.finditer(r'\? BLOCK [A-Z0-9]+ "?[A-Za-z0-9-]+"? \?', t)]
for pos, s in secs:
    print('  %7d  %s' % (pos, s))
print()

names = ['R-101', 'R-501', 'F-501', 'T-401', 'M-101', 'E-601', 'T-301']
for nm in names:
    m = re.search(r'\? BLOCK [A-Z0-9]+ "%s" \?' % re.escape(nm), t)
    if not m:
        print('--- %s : 未找到段 ---' % nm)
        continue
    nxt = t.find('? BLOCK', m.end())
    nxt2 = t.find('? ', m.end())
    end = nxt if nxt != -1 else nxt2
    seg = re.sub(r'\s+', ' ', t[m.start():end])
    print('--- %s (%d 字符) ---' % (nm, len(seg)))
    print(seg[:700])
    print()
