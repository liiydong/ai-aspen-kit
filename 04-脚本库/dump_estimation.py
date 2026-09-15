# -*- coding: utf-8 -*-
"""dump Properties\Estimation，试开"估算全部缺失参数"，Run 后看结果"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>'


def dump(node, path, indent, out, maxn=40, depth=0, maxdepth=3):
    if node is None:
        out.append('%s%s = None' % (indent, path)); return
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


msgs = []


class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            msgs.append(s)


doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
win32.WithEvents(doc, Sink)
doc.InitFromArchive2(os.path.join(BASE, 'NA_work.bkp'))
t = time.time()
while time.time() - t < 25:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)

out = []
out.append('##### \\Data\\Properties\\Estimation')
dump(doc.Tree.FindNode(r'\Data\Properties\Estimation'), 'Estimation', '  ', out, 40, 0, 3)
out.append('')
out.append('##### \\Data\\Properties\\Property Methods（前后）')
dump(doc.Tree.FindNode(r'\Data\Properties\Property Methods'), 'PM', '  ', out, 20, 0, 2)
print('\n'.join(out))

# 找 "Estimate all missing parameters" 对应字段
print()
print('--- 搜索 Estimation 下的开关字段 ---')
for p in [r'\Data\Properties\Estimation',
          r'\Data\Properties\Estimation\Input',
          r'\Data\Properties\Estimation\Output']:
    nd = doc.Tree.FindNode(p)
    if nd is None:
        print('  %s None' % p); continue
    try:
        for i in range(nd.Elements.Count):
            e = nd.Elements.Item(i)
            try:
                nm = e.Name
            except Exception:
                continue
            try:
                cc = e.Elements.Count
                print('  %s\\%s [%d]' % (p.split('\\')[-1], nm, cc))
                for j in range(min(cc, 12)):
                    ee = e.Elements.Item(j)
                    try:
                        print('       %s = %r' % (ee.Name, ee.Value))
                    except Exception:
                        pass
            except Exception:
                try:
                    print('  %s\\%s = %r' % (p.split('\\')[-1], nm, e.Value))
                except Exception:
                    pass
    except Exception as ex:
        print('  ' + str(ex)[:90])

print('DONE')
