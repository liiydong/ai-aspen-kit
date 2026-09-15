# -*- coding: utf-8 -*-
"""对比 DATABANKS 段与 NRTL 参数段原文（保留换行）"""
import re, os

def show(path, tag):
    print('=' * 70)
    print(tag, '|', os.path.basename(path), '|', os.path.getsize(path), 'bytes')
    print('=' * 70)
    t = open(path, encoding='utf-8', errors='ignore').read()

    # 1) DATABANKS 段
    # 段标记形式：\n DATABANKS \n  或  ? DATABANKS ?
    m = re.search(r'(\\\s*\n?DATABANKS.*?)(?=\n\\ ?[A-Z])', t, re.S)
    if not m:
        m = re.search(r'(DATABANKS.{0,400}?)(?=\n\\ ?[A-Z])', t, re.S)
    print('--- DATABANKS 段 ---')
    if m:
        seg = m.group(1)
        print('长度:', len(seg))
        print(repr(seg[:900]))
    else:
        # 直接搜 FILE-SYM-NAM
        k = t.find('FILE-SYM-NAM')
        print('未按段匹配；FILE-SYM-NAM 位置:', k)
        if k > 0:
            print(repr(t[max(0, k-300): k+500]))
    print()

    # 2) NRTL 参数段
    print('--- PARAMETERS BINARY (NRTL) 段 ---')
    i = t.find('PARAMETERS BINARY')
    print('位置:', i)
    if i > 0:
        print(repr(t[i:i+1400]))
    print()

    # 3) 统计 NRTL-1 表条数
    n = len(re.findall(r'NRTL-1', t))
    print('NRTL-1 出现次数:', n)
    # 所有形如 CID1/CID2 的行
    pairs = re.findall(r'CID1\s*=\s*"?([A-Za-z0-9_-]+)"?\s*', t)
    print('CID1 出现次数:', len(pairs), '唯一:', sorted(set(pairs))[:30])
    print()

for p, tag in [
    (r'D:\<化工工作区>\NA_work.bkp', '我们的定稿'),
    (r'D:\<化工工作区>\_probe\ref_multi.bkp', '参考（多效精馏）'),
]:
    if os.path.exists(p):
        show(p, tag)
    else:
        print('缺失:', p)
