# -*- coding: utf-8 -*-
"""dump 文件头 1600 字符（模型库清单 + 块实例注册表）"""
import sys
sys.stdout.reconfigure(encoding='utf-8')

t = open(r'D:\<化工工作区>\_probe\bipA2.bkp', encoding='utf-8', errors='ignore').read()
print('=== 头 1700 字符 ===')
print(t[:1700])
print()
print('=== repr 600-1000 ===')
print(repr(t[560:1000]))
