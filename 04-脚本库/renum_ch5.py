# -*- coding: utf-8 -*-
"""第5章：把新插入的水力学校核表定为 5-2，其余表号顺延；并修掉残留的旧投资数字"""
import re

P = r'D:\obsidian-vault\毕业设计\化学法版\05_第5章_主要设备工艺计算与选型.md'
t = open(P, encoding='utf-8').read()
orig = t

# 1) 我的新表临时改成占位
t = t.replace('表 5-2 精馏塔水力学校核', '表 5-X 精馏塔水力学校核', 1)
t = t.replace('结果见表 5-2。', '结果见表 5-X。', 1)

# 2) 旧表号 5-2..5-8 顺延 (降序处理避免连锁)
cnt = []
for n in range(8, 1, -1):
    old = '表 5-%d' % n
    new = '表 5-%d' % (n + 1)
    c = t.count(old)
    t = t.replace(old, new)
    cnt.append('%s->%s ×%d' % (old, new, c))

# 3) 占位改回 5-2
t = t.replace('表 5-X', '表 5-2')

# 4) 残留旧投资数字
fix = []
for a, b in [('5200 万元 vs', '4984 万元 vs'), ('5189 万元', '4984 万元'), ('5200 万元', '4984 万元')]:
    if a in t:
        fix.append('%s -> %s' % (a, b))
        t = t.replace(a, b)

open(P, 'w', encoding='utf-8').write(t)
print('表号顺延:', '; '.join(cnt))
print('投资数字修正:', fix)
print('文件长度 %d -> %d' % (len(orig), len(t)))
print()
print('=== 当前第5章表号 ===')
for m in re.finditer(r'^表 5-\d+.*$', t, re.M):
    print('  ', m.group(0)[:50])
print()
print('=== 正文引用 ===')
for m in re.finditer(r'.{0,18}表 5-\d+.{0,18}', t):
    s = m.group(0)
    if not s.startswith('表 5-'):
        print('  ', s.replace('\n', ' '))
