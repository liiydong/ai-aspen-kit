# -*- coding: utf-8 -*-
"""遍历 Components 子树，找组分实例的真实位置"""
import time
import win32com.client as win32

doc = win32.DispatchEx('Apwn.Document')
doc.InitFromArchive2(r'D:\<化工工作区>\_probe\na_latest.bkp')
time.sleep(4)

def kids(path):
    try:
        n = doc.Tree.FindNode(path)
        if n is None:
            return None
        return [(n.Elements.Item(i).Name, n.Elements.Item(i).Elements.Count)
                for i in range(n.Elements.Count)]
    except Exception as e:
        return 'ERR ' + str(e)[:70]

for p in [r'\Data\Components',
          r'\Data\Components\Specifications',
          r'\Data\Components\Specifications\Input',
          r'\Data\Components\Specifications\Input\Input',
          r'\Data\Components\Comp-Lists']:
    print('=' * 12, p)
    r = kids(p)
    if isinstance(r, str):
        print('  ', r)
    elif r is None:
        print('  None')
    else:
        for name, cnt in r[:20]:
            print(f'   {name}  [{cnt}]')
    print()
