# -*- coding: utf-8 -*-
"""加强版孤儿表清理：容忍表格之间夹着的引用/说明行"""
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

WINDOW = 60   # 在标题后多少行内查表

for P, titles in TARGETS.items():
    lines = open(P, encoding='utf-8').read().split('\n')
    # 从后往前处理，避免行号变化
    for title in titles:
        ti = None
        for i, ln in enumerate(lines):
            if ln.startswith(title):
                ti = i
                break
        if ti is None:
            LOG.append('!! 未找到 %s | %s' % (P.split('/')[-1], title)); continue
        blocks = []   # (start,end)
        i = ti + 1
        end = min(len(lines), ti + WINDOW)
        while i < end:
            if lines[i].startswith('|'):
                s = i
                while i < end and lines[i].startswith('|'):
                    i += 1
                blocks.append((s, i))
            else:
                i += 1
        if len(blocks) <= 1:
            LOG.append('OK  %s | %s : 仅 1 个表块' % (P.split('/')[-1], title[:22]))
            continue
        # 保留第一个；删除其余（连同其后的引用说明行）
        dels = []
        for (s, e) in blocks[1:]:
            k = e
            while k < end and (lines[k].startswith('>') or lines[k].strip() == ''):
                k += 1
            dels.append((s, k))
        for (s, e) in sorted(dels, reverse=True):
            # 回退吃掉前导空行
            while s - 1 > ti and lines[s - 1].strip() == '':
                s -= 1
            del lines[s:e]
        LOG.append('CLEAN %s | %s : 删除 %d 个孤儿表块' % (P.split('/')[-1], title[:22], len(dels)))
    open(P, 'w', encoding='utf-8').write('\n'.join(lines))

open(r'D:\<化工工作区>\_probe\clean2_log.txt', 'w', encoding='utf-8').write('\n'.join(LOG))
print('\n'.join(LOG))
