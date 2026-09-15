# -*- coding: utf-8 -*-
"""尝试各种添加组分的方式"""
import time
import win32com.client as win32

doc = win32.DispatchEx('Apwn.Document')
doc.InitFromArchive2(r'D:\<化工工作区>\_probe\na_from_apwz.bkp')
time.sleep(3)
print('已打开')

node = doc.Tree.FindNode(r'\Data\Components\Specifications\Input')
print('Input 子元素数:', node.Elements.Count)

tests = [
    ('Add()', lambda: node.Elements.Add()),
    ('Add("3-MP")', lambda: node.Elements.Add('3-MP')),
    ('Add("3-MP","3-MP")', lambda: node.Elements.Add('3-MP', '3-MP')),
    ('Add("3-MP","108-99-6")', lambda: node.Elements.Add('3-MP', '108-99-6')),
    ('Add2("3-MP")', lambda: node.Elements.Add2('3-MP')),
]
for name, fn in tests:
    try:
        r = fn()
        print(f'  {name}: 成功  返回={r}  现在子元素数={node.Elements.Count}')
        break
    except Exception as e:
        print(f'  {name}: 失败 {str(e)[:110]}')

print('\n--- 看看 Specifications 层 ---')
sp = doc.Tree.FindNode(r'\Data\Components\Specifications')
try:
    print('Specifications 子元素:', [sp.Elements.Item(i).Name for i in range(sp.Elements.Count)])
except Exception as e:
    print('失败:', str(e)[:120])

print('\n--- 看看 Input 的全部 16 个子节点 ---')
try:
    names = [node.Elements.Item(i).Name for i in range(node.Elements.Count)]
    print(names)
except Exception as e:
    print('失败:', str(e)[:120])
