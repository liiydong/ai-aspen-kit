# -*- coding: utf-8 -*-
"""定位：apw 能否加载 flowsheet；mod.bkp 的组分是否被接受"""
import time
import win32com.client as win32


def report(path, tag, method='InitFromArchive'):
    print('=' * 20, tag)
    doc = win32.DispatchEx('Apwn.Document')
    try:
        doc.SuppressDialogs = True
    except Exception:
        pass
    try:
        getattr(doc, method)(path)
    except Exception as e:
        print('  打开失败:', str(e)[:150])
        return
    time.sleep(3)
    ms = doc.Tree.FindNode(r'\Data\Properties\Molecular Structure')
    if ms is None:
        print('  组分节点: None')
    else:
        names = [ms.Elements.Item(i).Name for i in range(ms.Elements.Count)]
        print('  组分数 %d: %s' % (len(names), names))
    for p in [r'\Data\Streams', r'\Data\Blocks']:
        n = doc.Tree.FindNode(p)
        print('  %-16s %s' % (p, 'None' if n is None else '存在(子=%s)' % n.Elements.Count))
    try:
        doc.Close()
    except Exception:
        pass
    print()


report(r'D:\<化工工作区>\_probe\mod.bkp', 'mod.bkp (我注入6个组分)')
report(r'D:\<化工工作区>\_probe\na_main.apw', 'na_main.apw (二进制主文件)')
report(r'D:\<化工工作区>\_probe\base.bkp', 'base.bkp (用户原始12组分)')
