# -*- coding: utf-8 -*-
"""交付前自检：字数、乱码、表号连续性、关键指标一致性"""
import os, re, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
D = r'C:\Users\<用户名>\<项目原始资料>化学法版'
THESIS = ['00_封面与摘要.md', '01_第1章_绪论.md', '02_第2章_生产工艺流程设计与说明.md',
          '03_第3章_物料衡算.md', '04_第4章_热量衡算.md', '05_第5章_主要设备工艺计算与选型.md',
          '06_第6章_环境保护与安全生产.md', '07_结论参考文献与附录.md']
out = []
tot = 0
for fn in THESIS:
    t = open(os.path.join(D, fn), encoding='utf-8').read()
    cn = len(re.findall(r'[\u4e00-\u9fff]', t))
    tot += cn
    bad = re.findall(r'[\ufffd\u25a1]|Ã|â€', t)
    tbl = re.findall(r'表\s*(\d+)\s*[-–]\s*(\d+)', t)
    nums = {}
    for ch, n in tbl:
        nums.setdefault(ch, []).append(int(n))
    seqtxt = []
    for ch in sorted(nums):
        s = nums[ch]
        seqtxt.append('章%s:[%s]%s' % (ch, ','.join(map(str, s)),
                                       '' if s == list(range(1, len(s) + 1)) else ' <--不连续/重复'))
    out.append('%-42s 中文 %6d 字  乱码 %d  表号 %s' % (fn, cn, len(bad), ' | '.join(seqtxt)))
out.append('')
out.append('正文中文合计：%d 字' % tot)
out.append('')
out.append('--- 关键指标出现次数 ---')
KEYS = ['1411.6', '10164', '916.9', '83.26', '10163.8', '44000', '6.11 t/h', '3.9 t/h',
        '2193.6', '1822 kW', '1730', '2696.0', '70.72', '5353.5', '2657.4', '1254.5',
        '10221', '1419.5', '910.7', '83.73', '11.8 t/h', '933.8']
for k in KEYS:
    c = sum(open(os.path.join(D, f), encoding='utf-8').read().count(k) for f in THESIS)
    out.append('  %-12s %d' % (k, c))
print('\n'.join(out))
