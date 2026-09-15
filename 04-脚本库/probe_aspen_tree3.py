# -*- coding: utf-8 -*-
import time, sys
import win32com.client as win32

doc = win32.Dispatch('Apwn.Document')
doc.InitNew()
time.sleep(3)
print('InitNew OK')

paths = [r'\Data\Components', r'\Data\Components\Specifications\Input',
         r'\Data\Blocks', r'\Data\Streams', r'\Data\Setup\Specifications\Input']
for p in paths:
    try:
        n = doc.Tree.FindNode(p)
        print(f'\n{p}\n  Name={n.Name} 子元素={n.Elements.Count} Value={str(n.Value)[:40]}')
        for i in range(min(12, n.Elements.Count)):
            print('    ', i, n.Elements.Item(i).Name)
    except Exception as e:
        print(f'\n{p} 失败: {repr(e)[:120]}')

print('\n--- Add 签名试探 ---')
blk = doc.Tree.FindNode(r'\Data\Blocks')
for a in [('B1',), ('B1', 'RStoic'), ('B1', 'RStoic', 'reactors'),
          ('B1', 'RStoic', 'REACTORS'), ('B1', 'RStoic', 'AspenTech')]:
    try:
        blk.Elements.Add(*a)
        print('  OK Add', a, '-> Blocks 数 =', blk.Elements.Count)
        break
    except Exception as e:
        print('  FAIL Add', a, ':', str(e)[:120])

print('\n--- 组分写入路径试探 ---')
cand = [r'\Data\Components\Specifications\Input\COMPONENT_LIST',
        r'\Data\Components\Specifications\Input\Components',
        r'\Data\Components\Specifications\Input']
for p in cand:
    try:
        n = doc.Tree.FindNode(p)
        print(f'  {p} 存在，子元素={n.Elements.Count}')
    except Exception as e:
        print(f'  {p} 失败: {str(e)[:100]}')
print('done')
