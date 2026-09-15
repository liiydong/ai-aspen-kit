# -*- coding: utf-8 -*-
"""取本机权威数据库符号名：从健康模型 + Aspen 安装目录"""
import sys, re, os, glob
sys.stdout.reconfigure(encoding='utf-8')
L = []

# 1) 健康模型（多效精馏）的 DATABANKS 段全文
for p in [r'D:\<化工工作区>\_probe\ref_multi.bkp',
          r'D:\<化工工作区>\_probe\fresh_default.bkp']:
    if not os.path.exists(p):
        L.append('MISS %s' % p); continue
    t = open(p, encoding='utf-8', errors='ignore').read()
    flat = re.sub(r'[ \t]+', ' ', t)
    L.append('===== %s =====' % os.path.basename(p))
    m = re.search(r'\?\s*DATABANKS\s*\?', flat)
    if m:
        j = re.search(r'\n\?\s*[A-Z]', flat[m.end():])
        e = m.end() + (j.start() if j else 3000)
        L.append(flat[m.start():min(e, m.start() + 3000)])
    else:
        L.append('  无 DATABANKS 段')
    # 找 FILE-SYM-NAM / DEF-SYM-NAM
    for key in ['FILE-SYM-NAM', 'DEF-SYM-NAM']:
        for mm in re.finditer(key, flat):
            i = mm.start()
            L.append('  [%s] %s' % (key, flat[i:i + 400]))
    L.append('')

# 2) Aspen 安装目录里的数据库定义文件
cands = []
for root in [r'C:\Program Files\AspenTech\AprSystem V15.0',
             r'C:\Program Files\AspenTech\Aspen Plus V15.0',
             r'C:\ProgramData\AspenTech']:
    if not os.path.isdir(root):
        continue
    for pat in ['**/*databank*', '**/*.dbf', '**/bank*.txt', '**/*syn*.txt']:
        cands += glob.glob(os.path.join(root, pat), recursive=True)[:20]
L.append('===== 候选配置文件 (%d) =====' % len(cands))
for c in cands[:40]:
    L.append('  ' + c)

open(r'D:\<化工工作区>\_probe\db_names.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('\n'.join(L))
