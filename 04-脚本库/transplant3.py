# -*- coding: utf-8 -*-
"""移植 v3：正则支持带引号的组分名，把 15 个结构段全部搬到完整模型"""
import sys, re, os
sys.stdout.reconfigure(encoding='utf-8')

BASE = r'D:\<化工工作区>'
USER = os.path.join(BASE, 'NA_user.bkp')
FULL = os.path.join(BASE, 'NA_work.bkp')
OUT = os.path.join(BASE, 'NA_struct.bkp')

U = open(USER, encoding='utf-8', errors='ignore').read()
W = open(FULL, encoding='utf-8', errors='ignore').read()

# 组分名可选引号
PAT = re.compile(r'\? PROPERTIES "MOLEC-STRUCT" "?([A-Za-z0-9_\-]+)"?\s*\?')


def spans(x):
    ms = list(PAT.finditer(x))
    out = []
    for i, m in enumerate(ms):
        s = m.start()
        if i + 1 < len(ms):
            e = ms[i + 1].start()
        else:
            e = x.find('? PROPERTIES', m.end())
            if e < 0:
                e = len(x)
        out.append((m.group(1), s, e))
    return out


su = spans(U)
sw = spans(W)
print('用户文件段数: %d' % len(su))
print('完整模型段数: %d' % len(sw))

segs = {}
for name, s, e in su:
    segs[name] = U[s:e]

print()
print('%-8s %-8s %-6s %-6s %s' % ('组分', '长度', 'BONDS', 'FORMUL', '有无内容'))
for name, s, e in su:
    body = U[s:e]
    has = 'BONDS' in body or 'FORMUL' in body
    print('%-8s %-8d %-6d %-6d %s' % (name, len(body), body.count('BONDS'), body.count('FORMUL'), '有' if has else '空'))

# 从后往前替换
repl = 0
for i in range(len(sw) - 1, -1, -1):
    name, s, e = sw[i]
    if name not in segs:
        continue
    src = segs[name]
    if 'BONDS' not in src and 'FORMUL' not in src:
        continue      # 源也是空段，不覆盖
    W = W[:s] + src + W[e:]
    repl += 1
print()
print('替换 %d 个段' % repl)
open(OUT, 'w', encoding='utf-8', errors='ignore').write(W)
print('写出 %s  %d bytes' % (os.path.basename(OUT), os.path.getsize(OUT)))
V = open(OUT, encoding='utf-8', errors='ignore').read()
print('新文件 BONDS=%d FORMUL=%d' % (V.count('BONDS'), V.count('FORMUL')))
i = V.find('MOLEC-STRUCT" NAM')
print('NAM 样例:', repr(V[i:i + 240]))
