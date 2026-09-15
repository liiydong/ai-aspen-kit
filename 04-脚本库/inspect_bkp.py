# -*- coding: utf-8 -*-
import re

p = r'D:/<化工工作区>/NA-Chemical-10000t_修复.bkp'
t = open(p, encoding='utf-8', errors='ignore').read()
print('大小', len(t))
print('头部:', repr(t[:130]))
print()

m = re.search(r'RUN-CLASS = (\w+)', t)
print('RUN-CLASS =', m.group(1) if m else '?')
print()

j = t.find('? COMPONENTS MAIN ?')
k = t.find('? COMPONENTS "ADA/PCS"', j)
seg = re.sub(r'\s+', ' ', t[j + len('? COMPONENTS MAIN ?'):k])
ents = re.split(r'\s/\s', seg)
print('=== 组分共 %d 条 ===' % len(ents))
for i, e in enumerate(ents, 1):
    cid = re.search(r'CID = "?([^"\s]+)"?', e)
    db = re.search(r'DBNAME1 = "([^"]+)"', e)
    an = re.search(r'ANAME1 = "([^"]+)"', e)
    tp = re.search(r'TYPE = (\w+)', e)
    print('  %2d. %-12s DB=%-24s AN=%-14s %s' % (
        i, cid.group(1) if cid else '?',
        db.group(1) if db else '-', an.group(1) if an else '-',
        tp.group(1) if tp else ''))
print()

print('=== 段结构检查 ===')
for pat in [r'\? STREAM [A-Z0-9]+ ', r'\? BLOCK [A-Z0-9]+ ', r'FILE-SYM-NAM',
            r'\? TEAR \?', r'BLOCK BLKID', r'DEF-STREAM']:
    mm = re.findall(pat, t)
    print('  %-24s %d 个  %s' % (pat, len(mm), sorted(set(mm))[:6]))
print()
i = t.find('? DATABANKS ?')
print('=== DATABANKS 段 ===')
print(repr(t[i:i + 260]))
print()
print('=== 文件尾部 200 字符 ===')
print(repr(t[-200:]))
