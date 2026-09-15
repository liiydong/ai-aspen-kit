# -*- coding: utf-8 -*-
"""摸清 Aspen 树的参数写入路径"""
import time, sys
import win32com.client as win32

doc = win32.DispatchEx('Apwn.Document')
doc.InitFromArchive2(r'D:\<化工工作区>\_probe\na_from_apwz.bkp')
time.sleep(3)
print('文件已打开\n')

def dump(path, label, depth=0, maxdepth=1):
    try:
        n = doc.Tree.FindNode(path)
    except Exception as e:
        print(' ' * depth, f'{label}: ERR {str(e)[:70]}')
        return
    if n is None:
        print(' ' * depth, f'{label}: None')
        return
    try:
        cnt = n.Elements.Count
    except Exception:
        cnt = 0
    val = ''
    try:
        if n.Value is not None:
            val = str(n.Value)[:45]
    except Exception:
        pass
    print(' ' * depth, f'{label}  [{cnt}子]  {val}')
    if depth < maxdepth and cnt:
        for i in range(min(30, cnt)):
            el = n.Elements.Item(i)
            dump(path + '\\' + el.Name, el.Name, depth + 2, maxdepth)

for p, label in [(r'\Data\Streams\S-101\Input', '流股 S-101'),
                 (r'\Data\Blocks\R-101\Input', '模块 R-101'),
                 (r'\Data\Components\Specifications\Input', '组分'),
                 (r'\Data\Properties\Specifications\Input', '物性')]:
    print('=' * 15, label, p)
    dump(p, label, 0, 1)
    print()
