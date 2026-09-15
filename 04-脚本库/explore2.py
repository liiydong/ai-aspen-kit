# -*- coding: utf-8 -*-
"""探索：物性方法 / 反应 / 流股 / 模块 的参数写入路径"""
import time
import win32com.client as win32

P = r'D:\<化工工作区>\NA-Chemical-10000t_工作版.bkp'
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(P)
time.sleep(4)


def show(path, limit=60, withval=True):
    n = doc.Tree.FindNode(path)
    if n is None:
        print('  %-52s None' % path)
        return None
    try:
        c = n.Elements.Count
    except Exception:
        print('  %-52s (无子节点) val=%r' % (path, n.Value))
        return n
    print('  %s  [%d]' % (path, c))
    for i in range(min(c, limit)):
        ch = n.Elements.Item(i)
        try:
            v = repr(ch.Value)[:60] if withval else ''
        except Exception:
            v = '<err>'
        print('        - %-34s %s' % (ch.Name, v))
    return n


print('===== Properties =====')
show(r'\Data\Properties', withval=False)
show(r'\Data\Properties\Specifications\Input', 50)
print()
print('===== Reactions =====')
show(r'\Data\Reactions', withval=False)
show(r'\Data\Reactions\Reactions', withval=False)
el = doc.Tree.FindNode(r'\Data\Reactions\Reactions').Elements
print('    Reactions.Elements 属性:', [a for a in dir(el) if not a.startswith('_')][:18])
for desc, fn in [('Add()', lambda: el.Add()),
                 ('Add("R-1")', lambda: el.Add('R-1')),
                 ('Add("R-1","RSTOIC")', lambda: el.Add('R-1', 'RSTOIC'))]:
    try:
        fn()
        print('    %-24s 成功  count=%s' % (desc, el.Count))
    except Exception as e:
        print('    %-24s 失败 %s' % (desc, str(e)[:70]))
print()
print('===== Streams / S-101 =====')
show(r'\Data\Streams\S-101\Input', 40)
show(r'\Data\Streams\S-101\Input\FLOW', 20)
show(r'\Data\Streams\S-101\Input\FLOW\MIXED', 25)
print()
print('===== Blocks =====')
for b in ['M-101', 'E-101', 'R-101', 'T-201', 'T-401', 'R-501', 'F-501', 'E-601', 'T-301']:
    show(r'\Data\Blocks' + '\\' + b + r'\Input', 22, withval=False)
    print()
try:
    doc.Close()
except Exception:
    pass
