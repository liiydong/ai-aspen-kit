# -*- coding: utf-8 -*-
"""找组分与物性方法的正确树路径"""
import time
import win32com.client as win32

doc = win32.DispatchEx('Apwn.Document')
doc.InitFromArchive2(r'D:\<化工工作区>\_probe\na_from_apwz.bkp')
time.sleep(3)

def dump(path, depth=0, maxdepth=2):
    try:
        n = doc.Tree.FindNode(path)
    except Exception as e:
        print(' ' * depth + f'ERR {str(e)[:60]}')
        return
    if n is None:
        print(' ' * depth + 'None')
        return
    try:
        cnt = n.Elements.Count
    except Exception:
        cnt = 0
    try:
        val = '' if n.Value is None else str(n.Value)[:40]
    except Exception:
        val = ''
    print(' ' * depth + f'{path}  [{cnt}]  {val}')
    if depth < maxdepth * 2 and cnt:
        for i in range(min(25, cnt)):
            el = n.Elements.Item(i)
            dump(path + '\\' + el.Name, depth + 2, maxdepth)

for p in [r'\Data\Components', r'\Data\Properties', r'\Data\Setup']:
    print('=' * 20, p)
    dump(p, 0, 1)
    print()
