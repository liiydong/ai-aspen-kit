# -*- coding: utf-8 -*-
"""取完整错误信息 + 找 Heater/RadFrac 的真实字段名"""
import time
import win32com.client as win32

MID = r'D:\<化工工作区>\_probe\setup1.bkp'
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(MID)
time.sleep(4)

print('=== 完整错误：S-101 TEMP ===')
try:
    doc.Tree.FindNode(r'\Data\Streams\S-101\Input\TEMP').Value = 25.0
    print('  成功')
except Exception as e:
    print('  ', e)

print()
print('=== E-101 Input 中带 TEMP/PRES/SPEC 的字段 ===')
n = doc.Tree.FindNode(r'\Data\Blocks\E-101\Input')
for i in range(n.Elements.Count):
    ch = n.Elements.Item(i)
    if any(k in ch.Name.upper() for k in ['TEMP', 'PRES', 'SPEC', 'DUTY', 'VFRAC', 'OPT']):
        try:
            v = repr(ch.Value)[:40]
        except Exception:
            v = ''
        print('   %-24s %s' % (ch.Name, v))

print()
print('=== T-301 (Extract) 中带 NSTAGE/PRES/SPEC 的字段 ===')
n = doc.Tree.FindNode(r'\Data\Blocks\T-301\Input')
for i in range(n.Elements.Count):
    ch = n.Elements.Item(i)
    if any(k in ch.Name.upper() for k in ['NSTAGE', 'PRES', 'SPEC', 'NSTG', 'STAGE']):
        try:
            v = repr(ch.Value)[:40]
        except Exception:
            v = ''
        print('   %-24s %s' % (ch.Name, v))

print()
print('=== 分子结构节点 ===')
ms = doc.Tree.FindNode(r'\Data\Properties\Molecular Structure')
print('  子:', [ms.Elements.Item(i).Name for i in range(ms.Elements.Count)])
for c in ['NAM', '3-CP']:
    n2 = doc.Tree.FindNode(r'\Data\Properties\Molecular Structure' + '\\' + c)
    if n2 is None:
        print('  %s: None' % c)
        continue
    print('  %s 子节点:' % c, [n2.Elements.Item(i).Name for i in range(min(20, n2.Elements.Count))])
try:
    doc.Close()
except Exception:
    pass
