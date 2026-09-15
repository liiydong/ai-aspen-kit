# -*- coding: utf-8 -*-
"""检查各 bkp 是否存在重复的块段落"""
import sys, os, re
sys.stdout.reconfigure(encoding='utf-8')

FILES = [
    r'D:\<化工工作区>\_probe\bipA2.bkp',
    r'D:\<化工工作区>\NA-Chemical-10000t_v13_s.bkp',
    r'D:\<化工工作区>\NA-Chemical-10000t_修复.bkp',
    r'D:\<化工工作区>\NA-Chemical-10000t_模拟.bkp',
    r'D:\<化工工作区>\NA-Chemical-10000t_v17_s.bkp',
    r'D:\<化工工作区>\NA-Chemical-10000t_v19.bkp',
]
for p in FILES:
    if not os.path.exists(p):
        print('%-46s 不存在' % os.path.basename(p))
        continue
    t = open(p, encoding='utf-8', errors='ignore').read()
    nblk = len(re.findall(r'\?\s*BLOCK\s+RADFRAC', t))
    n401 = len(re.findall(r'BLOCK\s+RADFRAC\s*\n?\s*"T-401"', t))
    nstr = len(re.findall(r'\?\s*STREAM\s+MATERIAL\s*\n?\s*"S-111"', t))
    nprop = len(re.findall(r'\?\s*PROPERTIES\s+MAIN', t))
    ncl = len(re.findall(r'\?\s*COMPONENTS\s+MAIN', t))
    print('%-46s 大小=%-8d RADFRAC段=%-3d T-401段=%-3d S-111段=%-3d PROP_MAIN=%d COMP_MAIN=%d' % (
        os.path.basename(p), len(t), nblk, n401, nstr, nprop, ncl))
