# -*- coding: utf-8 -*-
import re

P = r'D:\<化工工作区>\_probe\panel_all.txt'
lines = open(P, encoding='utf-8').read().split('\n')
msgs = []
for ln in lines[1:]:
    m = re.match(r'\s*\d+ \| (.*)$', ln)
    if m:
        msgs.append(m.group(1))

print('消息数:', len(msgs))
print()
blocks = []
for i, m in enumerate(msgs):
    if re.search(r'\*\s*(WARNING|ERROR)', m):
        blk = [msgs[j] for j in range(i, min(i + 6, len(msgs)))]
        blocks.append((i, blk))
print('警告块数:', len(blocks))
for i, blk in blocks:
    print('--- @%d ---' % i)
    for b in blk:
        print('   ', b.replace('False', '').strip()[:160])
    print()

print('=== 最后 25 条 ===')
for m in msgs[-25:]:
    print('  ', m.replace('False', '').strip()[:170])
