# -*- coding: utf-8 -*-
"""把用户文件里的 MOLEC-STRUCT 段移植到完整模型，再试估算（换 UNIF-DMD）"""
import sys, re, os, time, shutil
sys.stdout.reconfigure(encoding='utf-8')

BASE = r'D:\<化工工作区>'
USER = os.path.join(BASE, 'NA_user.bkp')          # 用户回存（含结构，但流程图缺）
FULL = os.path.join(BASE, 'NA_work.bkp')          # 完整模型（缺结构）
OUT = os.path.join(BASE, 'NA_struct.bkp')

U = open(USER, encoding='utf-8', errors='ignore').read()
W = open(FULL, encoding='utf-8', errors='ignore').read()

# --- 提取用户文件里的 MOLEC-STRUCT 段 ---
segs = {}
for m in re.finditer(r'\? PROPERTIES "MOLEC-STRUCT" ([A-Za-z0-9_\-]+) \?', U):
    name = m.group(1)
    s = m.start()
    e = U.find('? PROPERTIES', s + 10)
    if e < 0:
        e = len(U)
    segs[name] = U[s:e]
print('用户文件中带结构的组分: %d 个' % len(segs))
for k in segs:
    body = segs[k]
    nb = body.count('BONDS')
    nf = body.count('FORMUL')
    print('  %-6s  BONDS=%d FORMUL=%d  长度=%d' % (k, nb, nf, len(body)))

# --- 逐个替换到完整模型 ---
cnt = 0
for name, seg in segs.items():
    pat = re.compile(r'\? PROPERTIES "MOLEC-STRUCT" %s \?' % re.escape(name))
    m = pat.search(W)
    if not m:
        print('  未在完整模型中找到段: %s' % name)
        continue
    s = m.start()
    e = W.find('? PROPERTIES', s + 10)
    if e < 0:
        e = len(W)
    W = W[:s] + seg + W[e:]
    cnt += 1
print('已替换 %d 个段' % cnt)

open(OUT, 'w', encoding='utf-8', errors='ignore').write(W)
print('写出 %s  %d bytes' % (os.path.basename(OUT), os.path.getsize(OUT)))

# 验证
V = open(OUT, encoding='utf-8', errors='ignore').read()
print('新文件 BONDS 计数:', V.count('BONDS'), '| FORMUL 计数:', V.count('FORMUL'))
i = V.find('MOLEC-STRUCT" TOL')
print('样例:', repr(V[i:i + 260]))
