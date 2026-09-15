# -*- coding: utf-8 -*-
"""探 RadFrac COL-SPECS 可用字段（找固定采出量的关键字）"""
import sys, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32

P = r'D:\<化工工作区>\_probe\bipA2.bkp'
doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
doc.InitFromArchive2(P)
time.sleep(4)


def names(path):
    n = doc.Tree.FindNode(path)
    if n is None:
        return None
    try:
        return [n.Elements.Item(k).Name for k in range(n.Elements.Count)]
    except Exception as ex:
        return 'ERR: %s' % ex


for p in [r'\Data\Blocks\T-401\Input',
          r'\Data\Blocks\T-401\Input\COL-SPECS',
          r'\Data\Blocks\T-401\Input\VIEW',
          ]:
    print(p, '->', names(p))
print()
# 列出 T-401\Input 下所有含 D / SPEC / RATE 的字段
n = doc.Tree.FindNode(r'\Data\Blocks\T-401\Input')
if n is not None:
    try:
        ch = [n.Elements.Item(k).Name for k in range(n.Elements.Count)]
        print('T-401\\Input 子节点(%d):' % len(ch))
        for c in ch:
            print('   ', c)
    except Exception as ex:
        print(ex)
try:
    doc.Close()
except Exception:
    pass
print('DONE')
