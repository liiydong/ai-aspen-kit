# -*- coding: utf-8 -*-
"""读分子结构内容 + UNIFAC Groups 两个节点"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>'


def dump(node, path, indent, out, maxn=30, depth=0, maxdepth=3):
    if node is None:
        out.append('%s%s = None' % (indent, path))
        return
    if depth > maxdepth:
        return
    try:
        c = node.Elements.Count
    except Exception:
        try:
            out.append('%s%s = %r' % (indent, path, node.Value))
        except Exception:
            out.append('%s%s = (?)' % (indent, path))
        return
    if c == 0:
        try:
            out.append('%s%s = %r (空)' % (indent, path, node.Value))
        except Exception:
            out.append('%s%s = (空表)' % (indent, path))
        return
    out.append('%s%s [%d]' % (indent, path, c))
    for i in range(min(c, maxn)):
        try:
            e = node.Elements.Item(i)
        except Exception:
            break
        if e is None:
            continue
        try:
            nm = e.Name
        except Exception:
            nm = 'item%d' % i
        dump(e, nm, indent + '   ', out, maxn, depth + 1, maxdepth)


doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
doc.InitFromArchive2(os.path.join(BASE, 'NA_work.bkp'))
t = time.time()
while time.time() - t < 25:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)

out = []
for comp in ['H2O', 'TOL', '3-MP', '3-CP', 'NAM']:
    out.append('##### Molecular Structure\\%s' % comp)
    dump(doc.Tree.FindNode(r'\Data\Properties\Molecular Structure\%s' % comp), comp, '  ', out, 20, 0, 2)
    out.append('')

for p in [r'\Data\Properties\Parameters\UNIFAC Groups',
          r'\Data\Properties\Parameters\UNIFAC Groups\Input',
          r'\Data\Properties\Parameters\UNIFAC Groups Binary',
          r'\Data\Properties\Parameters\UNIFAC Groups Binary\Input']:
    out.append('##### ' + p)
    dump(doc.Tree.FindNode(p), p.split('\\')[-1], '  ', out, 20, 0, 2)
    out.append('')

print('\n'.join(out))
try:
    doc.Close()
except Exception:
    pass
print('DONE')
