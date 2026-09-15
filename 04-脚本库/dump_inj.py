# -*- coding: utf-8 -*-
import re

p = r'D:\<化工工作区>\_probe\inj.bkp'
t = open(p, encoding='utf-8', errors='ignore').read()
i = t.find('PARAMNAME = NRTL')
j = t.find('PROPERTIES PARAMETERS BINARY', i + 40)
seg = t[i:j] if j > i else t[i:i+3000]
print('--- 注入后的 NRTL 段 ---')
print(seg)
print()
print('--- 该段关键计数 ---')
print('BPVAL =', seg.count('BPVAL'), ' UVAL =', seg.count('UVAL'))
print('反斜杠 =', seg.count('\\'))
