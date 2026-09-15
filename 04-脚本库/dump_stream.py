# -*- coding: utf-8 -*-
"""列出流股 Input 的全部字段名，找正确路径"""
import time
import win32com.client as win32

doc = win32.DispatchEx('Apwn.Document')
doc.InitFromArchive2(r'D:\<化工工作区>\_probe\na_from_apwz.bkp')
time.sleep(3)

n = doc.Tree.FindNode(r'\Data\Streams\S-101\Input')
names = [n.Elements.Item(i).Name for i in range(n.Elements.Count)]
print('S-101 Input 字段数:', len(names))
for x in names:
    print('   ', x)

print('\n--- 含 TEMP/PRES/FLOW 关键字的字段 ---')
for x in names:
    u = x.upper()
    if any(k in u for k in ('TEMP', 'PRES', 'FLOW', 'FRAC', 'VFRAC')):
        try:
            v = doc.Tree.FindNode(r'\Data\Streams\S-101\Input\\' + x).Value
        except Exception:
            v = '?'
        print(f'    {x:20s} = {v}')

print('\n--- 流股 S-101 的其它分支 ---')
s = doc.Tree.FindNode(r'\Data\Streams\S-101')
sub = [s.Elements.Item(i).Name for i in range(s.Elements.Count)]
print('  ', sub)
