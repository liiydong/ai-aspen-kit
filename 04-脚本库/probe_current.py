# -*- coding: utf-8 -*-
"""探测 COM 当前连接的 Aspen 实例与文件"""
import win32com.client as win32

doc = win32.Dispatch('Apwn.Document')
for attr in ('FullName', 'Name', 'Visible'):
    try:
        print(f'{attr} =', getattr(doc, attr))
    except Exception as e:
        print(f'{attr} 读取失败:', str(e)[:120])

print('\n--- 尝试访问树 ---')
try:
    n = doc.Tree.FindNode(r'\Data\Blocks')
    print('Blocks 子元素数 =', n.Elements.Count)
    for i in range(min(20, n.Elements.Count)):
        print('   ', i, n.Elements.Item(i).Name)
except Exception as e:
    print('Tree 访问失败:', str(e)[:200])

try:
    s = doc.Tree.FindNode(r'\Data\Streams')
    print('Streams 子元素数 =', s.Elements.Count)
    names = [s.Elements.Item(i).Name for i in range(min(40, s.Elements.Count))]
    print('   ', ', '.join(names))
except Exception as e:
    print('Streams 访问失败:', str(e)[:200])
