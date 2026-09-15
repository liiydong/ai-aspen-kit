# -*- coding: utf-8 -*-
"""论文全量一致性检查：乱码/重复/碎段/公式/表格/标题编号/引用"""
import glob, os, re, sys
sys.stdout.reconfigure(encoding='utf-8')
D = r'D:\obsidian-vault\毕业设计\化学法版'
FILES = sorted(glob.glob(os.path.join(D, '0[0-7]_*.md')))
L = []
def out(s=''):
    L.append(s)

CJK = re.compile(r'[\u4e00-\u9fff]')
out('=' * 70)
out('化学法版论文全量检查  %s' % __import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M'))
out('=' * 70)

total = 0
allcite = set()
refs = set()

for f in FILES:
    t = open(f, encoding='utf-8').read()
    name = os.path.basename(f)
    lines = t.split('\n')
    n = len(CJK.findall(t))
    total += n
    bad = t.count('\ufffd')
    out('')
    out('----- %s   中文字 %d  行数 %d  乱码 %d' % (name, n, len(lines), bad))

    # 1) 残留 LaTeX
    latex = re.findall(r'\\(?:frac|mathrm|times|cdot|rightarrow|sum|int|alpha|Delta|le|ge|approx)|\\\(|\\\)|\\\[|\\\]|\$\$', t)
    if latex:
        out('   [公式残留] %s' % set(latex))

    # 2) 粗体/斜体碎片（** 紧跟标点 或 空）
    frag = re.findall(r'\*\*\s*[\u4e00-\u9fff]{0,2}\s*\*\*', t)
    if frag:
        out('   [粗体碎片] %s' % frag[:6])

    # 3) 相邻重复串（≥12 中文字）
    for i in range(len(lines) - 1):
        a, b = lines[i].strip(), lines[i + 1].strip()
        if len(a) >= 12 and a == b:
            out('   [整行重复] 第%d行: %s' % (i + 1, a[:60]))
    # 行内重复：同一 15 字片段在一行内出现两次
    for i, ln in enumerate(lines):
        for m in re.finditer(r'([\u4e00-\u9fff]{12,}?)\1', ln):
            out('   [行内重复] 第%d行: %s' % (i + 1, m.group(1)[:50]))
            break

    # 4) 标题层级与编号
    heads = [(i + 1, m.group(1), m.group(2).strip())
             for i, ln in enumerate(lines)
             if (m := re.match(r'^(#{1,4})\s+(.*)$', ln))]
    nums = []
    for _, lvl, txt in heads:
        m = re.match(r'^(\d+(?:\.\d+)*)\s', txt)
        if m:
            nums.append((len(lvl), m.group(1), txt[:40]))
    # 检查同级编号连续
    bylevel = {}
    for lv, num, txt in nums:
        bylevel.setdefault(lv, []).append(num)
    for lv, arr in bylevel.items():
        bad2 = []
        for k in range(1, len(arr)):
            a = arr[k - 1].split('.')
            b = arr[k].split('.')
            if len(a) == len(b) and a[:-1] == b[:-1]:
                try:
                    if int(b[-1]) != int(a[-1]) + 1:
                        bad2.append('%s -> %s' % (arr[k - 1], arr[k]))
                except Exception:
                    pass
        if bad2:
            out('   [编号跳号 H%d] %s' % (lv, '; '.join(bad2[:8])))

    # 5) 表/图编号
    tabs = re.findall(r'表\s*(\d+)-(\d+)', t)
    figs = re.findall(r'图\s*(\d+)-(\d+)', t)
    for tag, arr in [('表', tabs), ('图', figs)]:
        if arr:
            ch = [x[0] for x in arr]
            seq = sorted(set(int(x[1]) for x in arr if x[0] == ch[0]))
            gaps = [i for i in range(1, max(seq) + 1 if seq else 1) if i not in seq]
            out('   [%s编号] 最大 %s-%d  缺号 %s' % (tag, ch[0], max(seq) if seq else 0,
                                                 gaps if gaps else '无'))

    # 6) 引用 [n]
    c = set(int(x) for x in re.findall(r'\[(\d{1,2})\]', t))
    allcite |= c

    # 7) 表格完整性：| 分隔行列数一致
    tbls, cur = [], []
    for i, ln in enumerate(lines):
        if ln.strip().startswith('|'):
            cur.append((i + 1, ln.count('|')))
        else:
            if len(cur) >= 2:
                tbls.append(cur)
            cur = []
    if len(cur) >= 2:
        tbls.append(cur)
    for tb in tbls:
        cnts = set(c for _, c in tb)
        if len(cnts) > 1:
            out('   [表格列数不一致] 起始行 %d, 列数分布 %s' % (tb[0][0], cnts))

# 参考文献清单
rf = os.path.join(D, '07_结论参考文献与附录.md')
if os.path.exists(rf):
    tt = open(rf, encoding='utf-8').read()
    for m in re.finditer(r'^\s*\[(\d{1,2})\]\s*(\S.{0,60})', tt, re.M):
        refs.add(int(m.group(1)))

out('')
out('=' * 70)
out('正文合计中文字符: %d' % total)
out('正文引用编号: %s' % sorted(allcite))
out('参考文献编号: %s' % sorted(refs))
out('引用但清单缺失: %s' % sorted(allcite - refs))
out('清单有但正文未引: %s' % sorted(refs - allcite))
open(r'D:\<化工工作区>\_probe\check_thesis.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('DONE')
