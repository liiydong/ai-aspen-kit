# -*- coding: utf-8 -*-
import glob, re, os

os.chdir(r'D:\obsidian-vault\毕业设计\化学法版')
tot = 0
for f in sorted(glob.glob('0[0-7]_*.md')):
    s = open(f, encoding='utf-8').read()
    bad = s.count('\ufffd')
    n = len(re.findall(r'[\u4e00-\u9fff]', s))
    tot += n
    print('%-48s 中文字=%-7d 乱码=%d' % (f, n, bad))
print('正文合计中文字符:', tot)
