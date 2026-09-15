# -*- coding: utf-8 -*-
import re

P = r'D:\<化工工作区>\NA-Chemical-10000t_BIP2.bkp'
t = open(P, encoding='utf-8', errors='ignore').read()
for k in ['FREE', 'NPHAS', 'NPHASE', 'FREEWATER', 'PHASE']:
    ms = re.findall(r'[A-Za-z_\-]*%s[A-Za-z_\-]*' % k, t[:120000])
    print('%-12s -> %s' % (k, sorted(set(ms))[:14]))

print()
m = re.search(r'FREE[^\\]{0,60}', t)
print('FREE 上下文:', repr(m.group(0)) if m else '无')
print()
# SETUP MAIN / SIM-OPTIONS 段落
i = t.find('? SETUP "SIM-OPTIONS"')
print('SIM-OPTIONS 段:', repr(t[i:i+300]) if i > 0 else '无')
print()
j = t.find('? SETUP MAIN ?')
print('SETUP MAIN 段:', repr(t[j:j+300]) if j > 0 else '无')
