# -*- coding: utf-8 -*-
"""对比：示例里 SEP 块的 FLOWSHEET 写法 vs 我们 T-301 的写法"""
import sys, os, re, glob
sys.stdout.reconfigure(encoding='utf-8')

# 我们文件
t = open(r'D:\<化工工作区>\NA-Chemical-10000t_v15.bkp', encoding='utf-8', errors='ignore').read()
i = t.find('"T-301"', t.find('FLOWSHEET GLOBAL'))
print('=== 我们 T-301 FLOWSHEET 原文 ===')
print(repr(t[i-90:i+230]))
print()

# 示例 Biodiesel
P = r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\Bulk Chemicals\Biodiesel Production from Vegetable Oil.bkp'
if not os.path.exists(P):
    c = glob.glob(r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\**\*Biodiesel*.bkp', recursive=True)
    P = c[0] if c else ''
print('示例文件:', P)
if P:
    s = open(P, encoding='latin-1', errors='ignore').read()
    for m in re.finditer(r'.{0,80}BLKTYPE = "SEP".{0,200}', s):
        print('---')
        print(repr(m.group()))
        break
    # 找 FLOWSHEET 段里的 SEP
    j = s.find('FLOWSHEET')
    k = s.find('BLKTYPE = "SEP"')
    print()
    print('=== 示例 SEP 块 FLOWSHEET ===')
    print(repr(s[k-120:k+260]) if k > 0 else '未找到')
