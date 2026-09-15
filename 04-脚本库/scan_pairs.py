# -*- coding: utf-8 -*-
"""在所有 Aspen 自带示例/数据包文件中，搜索含真实 UVAL 数值的二元参数对"""
import glob, os, re

ROOTS = [r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI',
         r'C:\ProgramData\AspenTech\APED V15.0']
FILES = []
for r in ROOTS:
    FILES += glob.glob(os.path.join(r, '**', '*.bkp'), recursive=True)
print('扫描 bkp 文件数:', len(FILES))

# 我们关心的组分别名
WANT = ['H2O', 'TOL', 'TOLUENE', 'NICOTINONITRILE', '4-PYRIDINENITRILE',
        'NICOTINIC-ACID-AMIDE', '4-METHYLPYRIDINE', '3-METHYLPYRIDINE',
        'WATER', 'METHYLPYRIDINE']

hits = []
for p in FILES:
    try:
        t = open(p, encoding='utf-8', errors='ignore').read()
    except Exception:
        continue
    if 'PARAMNAME = NRTL' not in t and 'PARAMNAME2 = NRTL' not in t:
        continue
    i = t.find('PARAMNAME = NRTL')
    j = t.find('? PROPERTIES', i + 20)
    seg = t[i:j] if j > i else t[i:i+80000]
    if 'UVAL' not in seg:
        continue
    # 提取所有 CID1/CID2 对
    pairs = set()
    for m in re.finditer(r'CID1\s*=\s*("?[\w\-\.]+"?)', seg):
        pass
    for m in re.finditer(r'CID1\s*=\s*("?[^"\s]+"?)\s+CID2\s*=\s*("?[^"\s]+"?)', re.sub(r'\s+', ' ', seg)):
        a = m.group(1).strip('"'); b = m.group(2).strip('"')
        pairs.add((a, b))
    if pairs:
        hits.append((p, len(seg), i, sorted(pairs)))

print('含真实数值的 NRTL 文件数:', len(hits))
for p, n, i, pairs in hits[:20]:
    print()
    print('==', os.path.basename(p), '(%d 文件, NRTL段 %d 字符, %d 对)' % (os.path.getsize(p), n, len(pairs)))
    rel = [x for x in pairs if any(w in (x[0] + x[1]).upper() for w in ['TOL', 'H2O', 'WATER', 'PYRIDINE', 'NICOT'])]
    print('   相关对:', rel[:20] if rel else '(无)')
    print('   全部:', pairs[:12])
