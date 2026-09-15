# -*- coding: utf-8 -*-
"""修正版：按位置切分提取 MOLEC-STRUCT 段（容忍折行）"""
import sys, re, os
sys.stdout.reconfigure(encoding='utf-8')

BASE = r'D:\<化工工作区>'
USER = os.path.join(BASE, 'NA_user.bkp')
FULL = os.path.join(BASE, 'NA_work.bkp')
OUT = os.path.join(BASE, 'NA_struct.bkp')

U = open(USER, encoding='utf-8', errors='ignore').read()
W = open(FULL, encoding='utf-8', errors='ignore').read()

PAT = re.compile(r'\? PROPERTIES "MOLEC-STRUCT" ([A-Za-z0-9_\-]+)\s*\?')
ms_u = list(PAT.finditer(U))
print('用户文件 MOLEC-STRUCT 段头: %d 个' % len(ms_u))

segs = {}
for i, m in enumerate(ms_u):
    name = m.group(1)
    s = m.start()
    if i + 1 < len(ms_u):
        e = ms_u[i + 1].start()
    else:
        e = U.find('? PROPERTIES', m.end())
        if e < 0:
            e = len(U)
    segs[name] = U[s:e]

print('提取到 %d 个段:' % len(segs))
for k, v in segs.items():
    print('  %-6s len=%-6d BONDS=%-3d FORMUL=%d' % (k, len(v), v.count('BONDS'), v.count('FORMUL')))

ms_w = list(PAT.finditer(W))
print()
print('完整模型 MOLEC-STRUCT 段头: %d 个' % len(ms_w))
names_w = [m.group(1) for m in ms_w]
print('  ', names_w)

# 逐个替换（从后往前，避免位移）
repl = 0
for i in range(len(ms_w) - 1, -1, -1):
    m = ms_w[i]
    name = m.group(1)
    if name not in segs:
        continue
    s = m.start()
    if i + 1 < len(ms_w):
        e = ms_w[i + 1].start()
    else:
        e = W.find('? PROPERTIES', m.end())
        if e < 0:
            e = len(W)
    W = W[:s] + segs[name] + W[e:]
    repl += 1
print()
print('替换 %d 个段' % repl)

open(OUT, 'w', encoding='utf-8', errors='ignore').write(W)
print('写出 %s  %d bytes' % (os.path.basename(OUT), os.path.getsize(OUT)))
V = open(OUT, encoding='utf-8', errors='ignore').read()
print('新文件 BONDS=%d  FORMUL=%d' % (V.count('BONDS'), V.count('FORMUL')))
i = V.find('MOLEC-STRUCT" TOL')
print('TOL 样例:', repr(V[i:i + 300]))
