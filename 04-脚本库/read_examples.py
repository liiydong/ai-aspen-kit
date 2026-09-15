# -*- coding: utf-8 -*-
"""读 Aspen 出厂示例的 DATABANKS 段 + NRTL 参数"""
import sys, re, os
sys.stdout.reconfigure(encoding='utf-8')

FILES = [
    r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\Bulk Chemicals\Distillation\3phase.bkp',
    r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\Batch Modeling\Batch Distillation\3phase.bkp',
    r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\Batch Modeling\Batch Distillation\WaterMethanol.bkp',
    r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\Bulk Chemicals\pfdtut.bkp',
]

for p in FILES:
    print('=' * 72)
    if not os.path.exists(p):
        print('缺失:', p)
        continue
    x = open(p, encoding='utf-8', errors='ignore').read()
    print(os.path.basename(p), '|', len(x), 'bytes')
    print('-' * 72)

    # DATABANKS 段（从 \ DATABANKS 到下一个 ? XXX ?）
    i = x.find('DATABANKS')
    print('--- DATABANKS 段 ---')
    if i >= 0:
        print(repr(x[i:i + 900]))
    else:
        print('未找到 DATABANKS')
    print()

    # 是否含 FILE-SYM-NAM
    print('含 FILE-SYM-NAM:', 'FILE-SYM-NAM' in x)

    # NRTL 参数条数
    j = x.find('PARAMNAME = NRTL')
    if j > 0:
        seg = x[j:j + 40000]
        k = seg.find('\n\\ ? ')
        if k > 0:
            seg = seg[:k]
        c1 = re.findall(r'CID1\s*=\s*"?([A-Za-z0-9_-]+)"?\s+CID2\s*=\s*"?([A-Za-z0-9_-]+)"?', seg)
        print('NRTL 二元对条数:', len(c1))
        srcs = sorted(set(re.findall(r'VAL1 = "([^"]+)"', seg)))
        print('数据来源 DATABANK:', srcs[:12])
        print('前 10 对:', ['%s-%s' % t for t in c1[:10]])
    else:
        print('无 NRTL 参数')
    print()
