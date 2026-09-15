# -*- coding: utf-8 -*-
"""列出 Aspen 树结构：Data 下各分支、Streams、一个流股的完整字段"""
import time, sys
import win32com.client as win32

P = sys.argv[1] if len(sys.argv) > 1 else r'D:\<化工工作区>\_probe\base.bkp'

doc = win32.DispatchEx('Apwn.Document')
doc.InitFromArchive2(P)
time.sleep(3)


def kids(node, limit=40):
    out = []
    try:
        n = node.Elements.Count
    except Exception as e:
        return ['<无Elements: %s>' % str(e)[:60]]
    for i in range(min(n, limit)):
        try:
            out.append(node.Elements.Item(i).Name)
        except Exception:
            out.append('<?>')
    return out


root = doc.Tree.FindNode(r'\Data')
print('\\Data 子节点 (%d):' % root.Elements.Count)
for k in kids(root):
    print('   ', k)
print()

st = doc.Tree.FindNode(r'\Data\Streams')
print('\\Data\\Streams 子节点数:', st.Elements.Count)
names = kids(st, 50)
print('   ', names)
print()

if names:
    s0 = doc.Tree.FindNode(r'\Data\Streams' + '\\' + names[0])
    print('流股 %s 子节点 (%d):' % (names[0], s0.Elements.Count))
    for k in kids(s0):
        print('   ', k)
    print()
    inp = doc.Tree.FindNode(r'\Data\Streams' + '\\' + names[0] + r'\Input')
    if inp is not None:
        print('  \\Input 子节点 (%d):' % inp.Elements.Count)
        for k in kids(inp, 60):
            print('     ', k)

try:
    doc.Close()
except Exception:
    pass
