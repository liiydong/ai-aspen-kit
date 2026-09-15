# -*- coding: utf-8 -*-
"""查流股规格字段并尝试设温压"""
import time
import win32com.client as win32

F = r'D:\<化工工作区>\NA-Chemical-10000t_模拟.bkp'
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(F)
time.sleep(4)

n = doc.Tree.FindNode(r'\Data\Streams\S-101\Input')
print('=== S-101 Input 里含 SPEC/FLASH/BASIS/TEMP/PRES 的字段 ===')
for i in range(n.Elements.Count):
    ch = n.Elements.Item(i)
    if any(k in ch.Name.upper() for k in ['SPEC', 'FLASH', 'BASIS', 'TEMP', 'PRES', 'FREE']):
        try:
            v = repr(ch.Value)[:45]
        except Exception:
            v = ''
        print('   %-20s %s' % (ch.Name, v))

print()
print('=== 尝试多种方式写 TEMP ===')
tries = [
    (r'\Data\Streams\S-101\Input\MIXED-SPEC', 'TP', 'MIXED-SPEC'),
    (r'\Data\Streams\S-101\Input\FLASH_FORM', 'PML', 'FLASH_FORM'),
    (r'\Data\Streams\S-101\Input\BASIS', 'MOLE-FLOW', 'BASIS'),
    (r'\Data\Streams\S-101\Input\TEMP', 25.0, 'TEMP(再次)'),
    (r'\Data\Streams\S-101\TEMP', 25.0, '\\Streams\\S-101\\TEMP'),
    (r'\Data\Streams\S-101\Output\TEMP', 25.0, 'Output\\TEMP'),
]
for path, val, tag in tries:
    try:
        nn = doc.Tree.FindNode(path)
        if nn is None:
            print('   %-26s 无节点' % tag)
            continue
        nn.Value = val
        print('   %-26s 成功 -> %r' % (tag, doc.Tree.FindNode(path).Value))
    except Exception as e:
        print('   %-26s 失败 %s' % (tag, str(e)[:80]))

print()
print('=== 属性方法是否生效 ===')
for p in [r'\Data\Properties\Specifications\Input\GOPSETNAME',
          r'\Data\Properties\Specifications\Input\GBASEOPSET']:
    nn = doc.Tree.FindNode(p)
    print('   %-22s %r' % (p.split('\\')[-1], nn.Value if nn is not None else None))
try:
    doc.Close()
except Exception:
    pass
