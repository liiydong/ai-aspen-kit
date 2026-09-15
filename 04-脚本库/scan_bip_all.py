# -*- coding: utf-8 -*-
"""扫描所有 bkp：DATABANKS 状态 + NRTL 二元对条数"""
import sys, re, os, glob
sys.stdout.reconfigure(encoding='utf-8')

BASE = r'D:\<化工工作区>'
files = []
for pat in ['*.bkp', '_probe/*.bkp']:
    files += glob.glob(os.path.join(BASE, pat))
files = sorted(set(files), key=lambda p: -os.path.getsize(p))

print('%-46s %10s %s' % ('文件', '大小', 'DATABANKS/NRTL对'))
print('-' * 100)
for p in files:
    try:
        x = open(p, encoding='utf-8', errors='ignore').read()
    except Exception:
        continue
    has_sym = 'FILE-SYM-NAM' in x
    i = x.find('PARAMNAME = NRTL')
    npair = 0
    pairs = []
    if i > 0:
        seg = x[i:i + 20000]
        # 只取到下一个段
        j = seg.find('\n\\ ? ')
        if j > 0:
            seg = seg[:j]
        c1 = re.findall(r'CID1\s*=\s*"?([A-Za-z0-9_-]+)"?\s+CID2\s*=\s*"?([A-Za-z0-9_-]+)"?', seg)
        npair = len(c1)
        pairs = ['%s-%s' % (a, b) for a, b in c1]
    print('%-46s %10d  sym=%-5s pairs=%-3d %s' % (
        os.path.basename(p), os.path.getsize(p), has_sym, npair,
        ', '.join(pairs[:8]) + (' ...' if len(pairs) > 8 else '')))
