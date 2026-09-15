# -*- coding: utf-8 -*-
"""细粒度重复/拼接扫描"""
import glob, os, re, sys
sys.stdout.reconfigure(encoding='utf-8')
D = r'D:\obsidian-vault\毕业设计\化学法版'
FILES = sorted(glob.glob(os.path.join(D, '0[0-7]_*.md')))
L = []
CJK = r'[\u4e00-\u9fff]'

for f in FILES:
    name = os.path.basename(f)
    lines = open(f, encoding='utf-8').read().split('\n')
    hits = []
    for i, ln in enumerate(lines):
        s = ln.strip()
        if not s:
            continue
        # a) 行内重复 >=6 中文字
        for m in re.finditer(r'(%s{6,}?)\1' % CJK, s):
            hits.append(('行内重复', i + 1, m.group(1)[:40]))
        # b) 词块紧邻重复：X...X 且间隔 <25 字（含英文）
        for m in re.finditer(r'([\u4e00-\u9fffA-Za-z0-9\-]{5,20})(?:\S{0,12})\1', s):
            hits.append(('近邻重复', i + 1, m.group(1)[:40]))
        # c) 句末标点后又出现同一句首
        # d) 缺标点拼接：中文字紧跟英文数字再跟中文字（可能是漏标点）
        for m in re.finditer(r'%s(?=[A-Z]-?\d)' % CJK, s):
            seg = s[max(0, m.start() - 12):m.start() + 16]
            if re.search(r'%s[A-Z]-?\d{2,3}%s' % (CJK, CJK), seg):
                hits.append(('疑似拼接', i + 1, seg))
    # e) 相邻行重复 >=8 字（一行是另一行前缀）
    for i in range(len(lines) - 1):
        a, b = lines[i].strip(), lines[i + 1].strip()
        if len(a) >= 10 and len(b) >= 10 and (a.startswith(b[:10]) or b.startswith(a[:10])) and a != b:
            hits.append(('相邻相似', i + 1, a[:36] + ' || ' + b[:36]))
    if hits:
        L.append('----- %s  (%d 处)' % (name, len(hits)))
        seen = set()
        for k, ln, s in hits:
            key = (k, s)
            if key in seen:
                continue
            seen.add(key)
            L.append('   [%s] 第%d行: %s' % (k, ln, s))
        L.append('')
open(r'D:\<化工工作区>\_probe\scan_dup.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('DONE %d' % len(L))
