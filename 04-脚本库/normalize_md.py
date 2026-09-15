# -*- coding: utf-8 -*-
"""Markdown 规范化 + 最后小修补"""
import sys, re, glob, os
sys.stdout.reconfigure(encoding='utf-8')
LOG = []
D = r'D:\obsidian-vault\毕业设计\化学法版'

# ---------- A) 小修补 ----------
P5 = os.path.join(D, '05_第5章_主要设备工艺计算与选型.md')
t = open(P5, encoding='utf-8').read()
old = ('- 理论板数 45 块，实际板数按 70% 板效率计算 = 64 块\n'
       '- 塔高：64 × 0.45 m + 2.5 m = 31.3 m\n'
       '- 因塔过高，采用两座串联的塔（T-403a 和 T-403b）每塔 32 块\n'
       '- 或采用丝网填料塔（CY 丝网），理论板高 0.3 m，45 块板只需 13.5 m\n'
       '- 本设计选用丝网填料方案：塔高 21 m（含上下封头）。')
new = ('- 理论板数 45 块。若按 70% 板效率、0.45 m 板间距的浮阀塔方案，实际板数与塔高分别达 64 块、31.3 m，'
       '已超出单塔经济高度；若拆为两座串联塔（T-403a/T-403b）则增加投资与占地\n'
       '- 故本设计改用丝网填料塔（CY 丝网）方案：理论板高 0.3 m，45 块理论板仅需 13.5 m 填料高\n'
       '- 最终塔高取 21 m（含上下封头与分布器空间）。')
if t.count(old) == 1:
    t = t.replace(old, new, 1); LOG.append('OK  ch5 5.2.2 方案表述')
else:
    LOG.append('!!  ch5 5.2.2 未命中 (%d)' % t.count(old))
open(P5, 'w', encoding='utf-8').write(t)

P2 = os.path.join(D, '02_第2章_生产工艺流程设计与说明.md')
t = open(P2, encoding='utf-8').read()
if '- 图 C-1 工艺总方块流程图（BFD）：' in t:
    t = t.replace('- 图 C-1 工艺总方块流程图（BFD）：', '- 图 C-1 工艺总方块流程图（BFD）：', 1)
    t = t.replace('- 图 C-2 工艺物料流程图（PFD）：', '- 图 C-2 工艺物料流程图（PFD）：', 1)
    LOG.append('OK  ch2 图名（无尾冒号需处理）')
# 去掉列表项尾部的裸冒号
t = re.sub(r'^(- 图 C-\d[^\n]*?)：\s*$', r'\1', t, flags=re.M)
open(P2, 'w', encoding='utf-8').write(t)
LOG.append('OK  ch2 图名尾冒号清理')

# ---------- B) 规范化：标题/引用行前后空行；压缩 3+ 连续空行 ----------
for f in sorted(glob.glob(os.path.join(D, '0[0-7]_*.md'))):
    lines = open(f, encoding='utf-8').read().split('\n')
    out = []
    for i, ln in enumerate(lines):
        if re.match(r'^#{1,4}\s', ln) and out and out[-1].strip() != '':
            out.append('')
        out.append(ln)
        if ln.startswith('>') :
            nxt = lines[i + 1] if i + 1 < len(lines) else ''
            if nxt.strip() != '' and not nxt.startswith('>'):
                out.append('')
    # 压缩连续空行
    res = []
    blank = 0
    for ln in out:
        if ln.strip() == '':
            blank += 1
            if blank > 2:
                continue
        else:
            blank = 0
        res.append(ln)
    txt = '\n'.join(res)
    open(f, 'w', encoding='utf-8').write(txt)
    LOG.append('OK  规范化 %s' % os.path.basename(f))

open(r'D:\<化工工作区>\_probe\norm_log.txt', 'w', encoding='utf-8').write('\n'.join(LOG))
print('\n'.join(LOG))
