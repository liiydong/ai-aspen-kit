# -*- coding: utf-8 -*-
import re, sys
sys.stdout.reconfigure(encoding='utf-8')
p = r'D:\<化工工作区>\NA-Chemical-10000t_最终定稿.bkp'
t = open(p, encoding='utf-8', errors='ignore').read()
L = []
L.append('===== T-403 全部出现位置 (%d) =====' % len(re.findall(r'T-403', t)))
for m in re.finditer(r'T-403', t):
    i = m.start()
    L.append('  pos %7d : %s' % (i, re.sub(r'\s+', ' ', t[max(0, i - 150):i + 200])))
    L.append('')
# 文件头注册表：找声明块数量的地方
L.append('===== 文件头 0-2000 字符 =====')
L.append(re.sub(r'[ \t]+', ' ', t[:2000]))
# RadFrac 与 SEP 段模板对比
L.append('')
L.append('===== T-402 RADFRAC 段（模板）=====')
m = re.search(r'\?\s*BLOCK\s+RADFRAC\s+"?T-402"?\s*\?', t)
if m:
    j = re.search(r'\n\?\s*BLOCK', t[m.end():])
    j = m.end() + (j.start() if j else 3000)
    L.append(t[m.start():min(j, m.start() + 4000)])
L.append('')
L.append('===== T-403 SEP 段（现状）=====')
m = re.search(r'\?\s*BLOCK\s+SEP\s+"?T-403"?\s*\?', t)
if m:
    j = re.search(r'\n\?\s*BLOCK', t[m.end():])
    j = m.end() + (j.start() if j else 2500)
    L.append(t[m.start():min(j, m.start() + 2500)])
open(r'D:\<化工工作区>\_probe\t403.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('DONE')
