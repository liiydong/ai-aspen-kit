# -*- coding: utf-8 -*-
"""第5章（投资合计）+ 第0章（摘要/目录）修订"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
LOG = []

def load(p):
    return open(p, encoding='utf-8').read()
def save(p, s):
    open(p, 'w', encoding='utf-8').write(s)

def mk(P, desc):
    global t
    t = load(P)
    LOG.append('=== %s' % desc)
def done(P):
    save(P, t)

def rep(old, new, desc):
    global t
    n = t.count(old)
    if n != 1:
        LOG.append('!! 失败(%d处): %s' % (n, desc)); return
    t = t.replace(old, new, 1); LOG.append('OK  %s' % desc)

# =============== 第5章 ===============
P5 = r'D:\obsidian-vault\毕业设计\化学法版\05_第5章_主要设备工艺计算与选型.md'
mk(P5, '第5章')
rep('（与附录 B 明细表同口径，合计 5189 万元）', '（与附录 B 明细表同口径，合计 4984 万元）', '5.7 引导句')
rep('| **合计** | **5189** | **100.0%** |', '| **合计** | **4984** | **100.0%** |', '5.7 表5-5 合计行')
n = t.count('5189')
if n:
    t = t.replace('5189', '4984')
    LOG.append('OK  其余 5189 替换 %d 处' % n)
done(P5)

# =============== 第0章 ===============
P0 = r'D:\obsidian-vault\毕业设计\化学法版\00_封面与摘要.md'
mk(P0, '第0章')
rep('3-甲基吡啶单耗 937 kg/t（99% 商品量）。',
    '3-甲基吡啶单耗 936.89 kg/t（99% 商品量）。', '摘要 单耗')
rep('Aspen Plus V15 模拟涵盖从氨氧化 R-101 到三效蒸发 E-601 的核心工段，结晶/离心/干燥因溶解度数据不足采用手算。',
    'Aspen Plus V15 建立了覆盖全流程 36 个单元、61 条流股的稳态模型（物性方法 NRTL），'
    '模拟结果 Terminal/Severe/Errors 全为 0、总物料平衡闭合偏差 0.000%，'
    '得到成品烟酰胺 10221 t/a（超过设计规模 2.2%）、3-甲基吡啶单耗 910.7 kg/t（低于设计值 2.8%）、总收率 83.73%。',
    '摘要 Aspen 范围')
rep('Aspen Plus V15 is employed to simulate R-101 through E-601; crystallization, centrifugation and drying are calculated manually due to lack of solubility data.',
    'A steady-state Aspen Plus V15 model covering the whole flowsheet (36 units, 61 streams, NRTL) is established; '
    'the run converges with zero terminal/severe errors and a closed overall mass balance (deviation 0.000%), '
    'yielding 10 221 t/a of niacinamide (2.2% above the design capacity) at a 3-methylpyridine consumption of 910.7 kg/t.',
    'Abstract Aspen 范围')

# 目录：补 附录 D
rep('附录 A 主要物料物性表\n附录 B 设备一览表\n附录 C 工艺流程图',
    '附录 A 主要物料物性表\n附录 B 设备一览表\n附录 C 工艺流程图\n附录 D 缩写与符号对照表',
    '目录补附录D')
# 目录：第3章 3.4 名称对齐、第4章补 4.5
rep('3.4 后处理工段（结晶、离心、干燥）', '3.4 后处理工段', '目录 3.4')
done(P0)

open(r'D:\<化工工作区>\_probe\patch_ch50_log.txt', 'w', encoding='utf-8').write('\n'.join(LOG))
print('\n'.join(LOG))
