# -*- coding: utf-8 -*-
"""清理 rep_block 遗留的孤儿旧表"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
LOG = []

TARGETS = {
    'D:/obsidian-vault/毕业设计/化学法版/03_第3章_物料衡算.md': [
        '表 3-2 年产 10000 吨烟酰胺装置物料平衡总表',
        '表 3-3 关键进料物流的物性参数',
        '表 3-10 公用工程消耗汇总',
    ],
    'D:/obsidian-vault/毕业设计/化学法版/04_第4章_热量衡算.md': [
        '表 4-4 精馏塔热负荷汇总（塔顶温度按工业实测值对齐[1]）',
        '表 4-5 Aspen 模拟全塔结果',
        '表 4-6 关键流股模拟值与设计值对照',
        '表 4-9 全流程热负荷与公用工程消耗汇总',
        '表 4-10 化学法与酶法的热量消耗对比',
    ],
    'D:/obsidian-vault/毕业设计/化学法版/05_第5章_主要设备工艺计算与选型.md': [
        '表 5-1 四座精馏塔工艺参数',
        '表 5-2 主要换热器一览',
    ],
    'D:/obsidian-vault/毕业设计/化学法版/07_结论参考文献与附录.md': [
        '## 附录 B 设备一览表',
    ],
}

def block_spans(text, start):
    """从 start 起，返回连续表格块（允许中间空行）的 (s,e) 列表；遇到非表格非空行停止"""
    spans = []
    i = start
    n = len(text)
    while i < n:
        # 跳过空行
        j = i
        while j < n and text[j] in '\n\r':
            j += 1
        if j >= n or text[j] != '|':
            break
        # 读表格块
        k = j
        last = j
        while k < n:
            m = re.match(r'\|[^\n]*\n?', text[k:])
            if not m:
                break
            k += m.end()
            last = k
        spans.append((j, last))
        i = last
    return spans

for P, titles in TARGETS.items():
    t = open(P, encoding='utf-8').read()
    changes = 0
    for title in titles:
        i = t.find(title)
        if i < 0:
            LOG.append('!! %s 未找到: %s' % (P.split('/')[-1], title)); continue
        j = t.find('\n', i)
        spans = block_spans(t, j + 1)
        if len(spans) <= 1:
            LOG.append('OK  %s | %s : 无孤儿表' % (P.split('/')[-1], title[:24]))
            continue
        # 保留第一个，删除其余
        del_from = spans[1][0]
        del_to = spans[-1][1]
        # 向前吃掉多余空行
        keep_end = spans[0][1]
        seg = t[keep_end:del_from]
        t = t[:keep_end] + seg.rstrip('\n') + '\n\n' + t[del_to:]
        changes += 1
        LOG.append('CLEAN %s | %s : 删除 %d 个孤儿表块' % (P.split('/')[-1], title[:24], len(spans) - 1))
    if changes:
        open(P, 'w', encoding='utf-8').write(t)

open(r'D:\<化工工作区>\_probe\clean_orphan_log.txt', 'w', encoding='utf-8').write('\n'.join(LOG))
print('\n'.join(LOG))
