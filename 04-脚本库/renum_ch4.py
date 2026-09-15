# -*- coding: utf-8 -*-
"""第4章表号：按位置顺序重编 + 引用就近映射"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
P = r'D:\obsidian-vault\毕业设计\化学法版\04_第4章_热量衡算.md'
lines = open(P, encoding='utf-8').read().split('\n')
LOG = []

# 1) 收集所有 表 4-X 出现（含行内位置）
occ = []   # (abs_pos, lineno, old_num, is_title)
pos = 0
titles = []   # (abs_pos, new_num)
new_i = 0
for i, ln in enumerate(lines):
    for m in re.finditer(r'表\s*4-(\d+)', ln):
        is_title = ln.strip().startswith('表')
        occ.append([pos + m.start(), i, int(m.group(1)), is_title, m.group(0)])
    pos += len(ln) + 1

# 2) 先给表题编号
for o in occ:
    if o[3]:
        new_i += 1
        o.append(new_i)
        titles.append((o[0], o[2], new_i))
LOG.append('表题数 = %d' % new_i)
LOG.append('表题顺序: %s' % ' '.join('4-%d->4-%d' % (o[2], o[5]) for o in occ if o[3]))

# 3) 引用：就近的同号表题
for o in occ:
    if o[3]:
        continue
    cands = [x for x in titles if x[1] == o[2]]
    if not cands:
        o.append(o[2])
        LOG.append('!! 引用无目标: 4-%d (行%d)' % (o[2], o[1] + 1))
        continue
    best = min(cands, key=lambda x: abs(x[0] - o[0]))
    o.append(best[2])

# 4) 重写：从后往前替换
for o in sorted(occ, key=lambda x: -x[0]):
    p, lineno, old, is_t, txt, new = o
    lines[lineno] = lines[lineno][:0] + lines[lineno]   # noop
for o in sorted(occ, key=lambda x: -x[0]):
    p, lineno, old, is_t, txt, new = o
    ln = lines[lineno]
    # 找到该行内第几次出现需要替换（用全局偏移换算）
    # 简化：该行内所有该 old 的出现都替换为 new（同一行内同号引用与表题一致）
    ln2 = re.sub(r'表\s*4-%d(?!\d)' % old, '表 4-%d' % new, ln)
    lines[lineno] = ln2

open(P, 'w', encoding='utf-8').write('\n'.join(lines))
open(r'D:\<化工工作区>\_probe\renum_ch4_log.txt', 'w', encoding='utf-8').write('\n'.join(LOG))
print('\n'.join(LOG))
