# -*- coding: utf-8 -*-
"""在参考模型上穷举写流股温压的正确姿势"""
import time
import win32com.client as win32

REF = r'D:\<化工工作区>\_probe\ref_multi.bkp'
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(REF)
time.sleep(4)

print('=== 0402（进料流）Input 字段 ===', flush=True)
n = doc.Tree.FindNode(r'\Data\Streams\0402\Input')
if n is None:
    print('  不存在')
else:
    for i in range(n.Elements.Count):
        ch = n.Elements.Item(i)
        if any(k in ch.Name.upper() for k in ['TEMP', 'PRES', 'BASIS', 'SPEC', 'FLASH', 'TOTAL']):
            try:
                v = repr(ch.Value)[:40]
            except Exception:
                v = ''
            print('   %-18s %s' % (ch.Name, v), flush=True)

print()
print('=== 0402 写 TEMP 各种姿势 ===', flush=True)
TEST = [
    (r'\Data\Streams\0402\Input\TEMP', 25.0, 'float 25.0'),
    (r'\Data\Streams\0402\Input\TEMP', '25', 'str "25"'),
    (r'\Data\Streams\0402\Input\TEMP', '25 C', 'str "25 C"'),
    (r'\Data\Streams\0402\Input\TEMP', 25, 'int 25'),
]
for p, v, tag in TEST:
    try:
        doc.Tree.FindNode(p).Value = v
        print('   %-14s 成功 -> %r' % (tag, doc.Tree.FindNode(p).Value), flush=True)
        break
    except Exception as e:
        print('   %-14s 失败 %s' % (tag, str(e)[:75]), flush=True)

print()
print('=== 0402 写组成（对照）===', flush=True)
try:
    doc.Tree.FindNode(r'\Data\Streams\0402\Input\FLOW\MIXED\C7H8').Value = 1.0
    print('   组成写入成功 ->', doc.Tree.FindNode(r'\Data\Streams\0402\Input\FLOW\MIXED\C7H8').Value, flush=True)
except Exception as e:
    print('   组成写入失败', str(e)[:80], flush=True)

print()
print('=== 再试：先设 MIXED_SPEC 再写 TEMP ===', flush=True)
for p, v in [(r'\Data\Streams\0402\Input\MIXED_SPEC', 'TP'),
             (r'\Data\Streams\0402\Input\TEMP', 25.0)]:
    try:
        doc.Tree.FindNode(p).Value = v
        print('   %-40s 成功' % p.split('\\')[-1], flush=True)
    except Exception as e:
        print('   %-40s 失败 %s' % (p.split('\\')[-1], str(e)[:60]), flush=True)

print()
print('=== 试 Reinit 后再写 ===', flush=True)
try:
    doc.Reinit()
    time.sleep(2)
    doc.Tree.FindNode(r'\Data\Streams\0402\Input\TEMP').Value = 25.0
    print('   Reinit 后写入成功!', flush=True)
except Exception as e:
    print('   Reinit 后仍失败', str(e)[:80], flush=True)
try:
    doc.Close()
except Exception:
    pass
