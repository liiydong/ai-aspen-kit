# -*- coding: utf-8 -*-
"""检查 v16 中 T-301 段与紧随其后的段"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')

t = open(r'D:\<化工工作区>\NA-Chemical-10000t_v16.bkp', encoding='utf-8', errors='ignore').read()
i = t.find('? BLOCK SEP "T-301"')
print('=== T-301 段原文 ===')
print(repr(t[i:i + 1000]))
print()
j = t.find('\\ ', i + 1000)
# 找紧随其后的 '? BLOCK'
k = re.search(r'\? BLOCK', t[i + 100:])
print('=== 其后第一个 ? BLOCK 位置 ===', i + 100 + k.start() if k else None)
print(repr(t[i + 900:i + 1350]))
