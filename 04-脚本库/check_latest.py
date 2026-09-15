# -*- coding: utf-8 -*-
"""检查用户最新文件的状态：组分、拓扑、可写性"""
import time, sys
import win32com.client as win32

P = r'D:\<化工工作区>\_probe\na_latest.bkp'
doc = win32.DispatchEx('Apwn.Document')
print('打开中 ...')
sys.stdout.flush()
opened = False
for m in ('InitFromArchive2', 'InitFromArchive'):
    try:
        getattr(doc, m)(P)
        print(f'{m} 打开成功')
        opened = True
        break
    except Exception as e:
        print(f'{m} 失败: {str(e)[:110]}')
if not opened:
    sys.exit(1)

time.sleep(4)

c = doc.Tree.FindNode(r'\Data\Components\Specifications\Input')
print('\n组分节点元素数:', c.Elements.Count)
names = [c.Elements.Item(i).Name for i in range(c.Elements.Count)]
print('子节点:', names)

b = doc.Tree.FindNode(r'\Data\Blocks')
s = doc.Tree.FindNode(r'\Data\Streams')
print('\n模块数:', b.Elements.Count, '| 流股数:', s.Elements.Count)

print('\n--- 测试流股写入（关键验证）---')
p = r'\Data\Streams\S-101\Input\TEMP'
try:
    doc.Tree.FindNode(p).Value = 25.0
    print('  写入成功，当前值 =', doc.Tree.FindNode(p).Value)
except Exception as e:
    print('  写入失败:', str(e)[:150])

print('\n--- 物性方法 ---')
for pp in (r'\Data\Properties\Specifications\Input\GLOBAL\METHOD',
           r'\Data\Properties\Specifications\Input\METHOD',
           r'\Data\Properties\Property Methods\Input\GLOBAL\METHOD'):
    try:
        v = doc.Tree.FindNode(pp).Value
        print(f'  {pp} = {v}')
    except Exception as e:
        print(f'  {pp}: {str(e)[:70]}')
