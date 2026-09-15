# -*- coding: utf-8 -*-
"""找分子结构节点 + 列 Engine 方法 + 列 Properties 子节点"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>'


def list_children(doc, path, out, limit=60):
    nd = doc.Tree.FindNode(path)
    if nd is None:
        out.append('  %s -> None' % path)
        return
    try:
        c = nd.Elements.Count
    except Exception:
        out.append('  %s -> leaf val=%r' % (path, nd.Value))
        return
    out.append('  %s  [%d]' % (path, c))
    for i in range(min(c, limit)):
        try:
            e = nd.Elements.Item(i)
        except Exception:
            break
        if e is None:
            continue
        try:
            nm = e.Name
        except Exception:
            nm = 'item%d' % i
        try:
            cc = e.Elements.Count
        except Exception:
            cc = '-'
        out.append('      %-30s rows=%s' % (nm, cc))


doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
doc.InitFromArchive2(os.path.join(BASE, 'NA_work.bkp'))
t = time.time()
while time.time() - t < 25:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)

out = []
out.append('##### \\Data\\Properties 子节点')
list_children(doc, r'\Data\Properties', out, 40)
out.append('')
for p in [r'\Data\Properties\Molecular Structure',
          r'\Data\Properties\MOLEC-STRUCT',
          r'\Data\Properties\Molecule Structure',
          r'\Data\Properties\Structure']:
    out.append('##### ' + p)
    list_children(doc, p, out, 12)
out.append('')
out.append('##### \\Data\\Components\\UNIFAC-Groups 完整')
list_children(doc, r'\Data\Components\UNIFAC-Groups', out, 20)
out.append('')
out.append('##### \\Data 顶层（再确认）')
list_children(doc, r'\Data', out, 30)

# Engine 方法
out.append('')
out.append('##### doc.Engine 方法')
try:
    for n in dir(doc.Engine):
        if not n.startswith('_'):
            out.append('    ' + n)
except Exception as e:
    out.append('  ' + str(e)[:80])

# Tree 方法
out.append('')
out.append('##### doc.Tree 方法')
try:
    for n in dir(doc.Tree):
        if not n.startswith('_'):
            out.append('    ' + n)
except Exception as e:
    out.append('  ' + str(e)[:80])

print('\n'.join(out))
try:
    doc.Close()
except Exception:
    pass
print('DONE')
