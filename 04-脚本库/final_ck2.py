# -*- coding: utf-8 -*-
import glob, re, os

os.chdir(r'D:\obsidian-vault\毕业设计\化学法版')
tot = 0
print('%-46s %8s %6s' % ('文件', '中文字', '乱码'))
for f in sorted(glob.glob('0[0-7]_*.md')):
    s = open(f, encoding='utf-8').read()
    n = len(re.findall(r'[\u4e00-\u9fff]', s))
    tot += n
    print('%-46s %8d %6d' % (f, n, s.count('\ufffd')))
print('正文合计:', tot)

print()
print('=== 各章表号连续性 ===')
for f in sorted(glob.glob('0[0-7]_*.md')):
    s = open(f, encoding='utf-8').read()
    ch = re.search(r'0(\d)_', f).group(1)
    caps = re.findall(r'^表 %s-(\d+)\s' % ch, s, re.M)
    if caps:
        nums = [int(x) for x in caps]
        ok = nums == sorted(nums) and (nums == list(range(1, len(nums) + 1)) or len(set(nums)) == len(nums))
        print('  第%s章: %s  %s' % (ch, nums, 'OK' if ok else '**不连续/重号**'))

print()
print('=== 公式残留检查（\\frac / \\sqrt 之外的 LaTeX 命令）===')
for f in sorted(glob.glob('0[0-7]_*.md')):
    s = open(f, encoding='utf-8').read()
    bad = re.findall(r'\\[a-zA-Z]+', s)
    from collections import Counter
    c = Counter(bad)
    extra = {k: v for k, v in c.items() if k not in
             ('\\frac', '\\sqrt', '\\times', '\\mathrm', '\\Delta', '\\rho',
              '\\pi', '\\cdot', '\\approx', '\\leq', '\\geq', '\\right', '\\left')}
    if extra:
        print(' ', f, extra)
