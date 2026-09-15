# -*- coding: utf-8 -*-
"""生成"只修运行类型"的干净文件，并验证"""
import time, subprocess
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\x_NA-Chemical-10000t.bkp.backup'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_修复.bkp'

text = open(SRC, encoding='utf-8', errors='ignore').read()
n = text.count('RUN-CLASS = PROP')
text = text.replace('RUN-CLASS = PROP', 'RUN-CLASS = FLOWSHEET')
open(OUT, 'w', encoding='utf-8', errors='ignore').write(text)
print('修复: RUN-CLASS = PROP -> FLOWSHEET  (%d 处)' % n, flush=True)
print('输出:', OUT, flush=True)
print()

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(OUT)
time.sleep(4)
d = doc.Tree.FindNode(r'\Data')
bl = doc.Tree.FindNode(r'\Data\Blocks')
st = doc.Tree.FindNode(r'\Data\Streams')
ms = doc.Tree.FindNode(r'\Data\Properties\Molecular Structure')
print('验证结果：', flush=True)
print('  \\Data 子节点  :', d.Elements.Count if d else 0, flush=True)
print('  模块 (Blocks) :', bl.Elements.Count if bl else 0, flush=True)
print('  流股 (Streams):', st.Elements.Count if st else 0, flush=True)
print('  组分          :', ms.Elements.Count if ms else 0, flush=True)
if bl and bl.Elements.Count:
    print('  模块清单      :', [bl.Elements.Item(i).Name for i in range(bl.Elements.Count)], flush=True)
try:
    doc.Close()
except Exception:
    pass
