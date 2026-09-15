# -*- coding: utf-8 -*-
"""细探 Aspen 树路径与 Elements.Add 签名"""
import win32com.client as win32

doc = win32.Dispatch('Apwn.Document')

paths = [r'\Data', r'\Data\Components', r'\Data\Components\Specifications',
         r'\Data\Components\Specifications\Input', r'\Data\Blocks', r'\Data\Streams']
for p in paths:
    try:
        n = doc.Tree.FindNode(p)
        print(f'{p}  ->  Name={n.Name}  子元素={n.Elements.Count}  Value={str(n.Value)[:50]}')
        for i in range(min(10, n.Elements.Count)):
            print('      ', i, n.Elements.Item(i).Name)
    except Exception as e:
        print(f'{p}  失败: {repr(e)[:110]}')

print('\n--- 试探 Elements.Add 的调用形式 ---')
blk = doc.Tree.FindNode(r'\Data\Blocks')
print('Blocks.Elements 类型:', type(blk.Elements))
tests = [('B1',), ('B1', 'RStoic'), ('B1', 'RStoic', 'reactors'),
         ('B1', 'RStoic', 'REACTORS'), ('RStoic',)]
for a in tests:
    try:
        blk.Elements.Add(*a)
        print('  Add', a, '成功；Blocks 数量 =', blk.Elements.Count)
        break
    except Exception as e:
        print('  Add', a, '失败:', str(e)[:110])

print('\n--- 用 gencache 取类型信息 ---')
try:
    import win32com.client.gencache as gc
    d = gc.EnsureDispatch('Apwn.Document')
    print('gencache 成功:', type(d))
except Exception as e:
    print('gencache 失败:', repr(e)[:150])
