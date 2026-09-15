# -*- coding: utf-8 -*-
import re, sys
sys.stdout.reconfigure(encoding='utf-8')
L = []
for ch in ['03_第3章_物料衡算', '04_第4章_热量衡算']:
    p = r'D:\obsidian-vault\毕业设计\化学法版\%s.md' % ch
    lines = open(p, encoding='utf-8').read().split('\n')
    L.append('===== %s =====' % ch)
    title_seq = []
    for i, ln in enumerate(lines):
        for m in re.finditer(r'表\s*(\d)-(\d+)', ln):
            is_title = ln.strip().startswith('表')
            tag = 'TITLE ' if is_title else 'ref   '
            L.append('  %s 行%4d  表%s-%s   | %s' % (tag, i + 1, m.group(1), m.group(2), ln.strip()[:80]))
            if is_title:
                title_seq.append('%s-%s' % (m.group(1), m.group(2)))
    L.append('  表题出现顺序: %s' % ' '.join(title_seq))
    L.append('')
# 跨章引用检查
L.append('===== 其他章引用 表3-/表4- =====')
import glob, os
for f in sorted(glob.glob(r'D:\obsidian-vault\毕业设计\化学法版\0*.md')):
    b = os.path.basename(f)
    if b.startswith('03_') or b.startswith('04_'):
        continue
    for i, ln in enumerate(open(f, encoding='utf-8').read().split('\n')):
        for m in re.finditer(r'表\s*[34]-\d+', ln):
            L.append('  %s 行%d: %s' % (b, i + 1, ln.strip()[:90]))
open(r'D:\<化工工作区>\_probe\tabrefs.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('DONE')
