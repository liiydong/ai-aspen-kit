# -*- coding: utf-8 -*-
"""深挖 Pure Databanks / Other Databanks / Components UNIFAC-Groups / NRTL 数据行"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>'


def walk(node, path, depth, maxdepth, out, limit=25):
    if depth > maxdepth:
        return
    try:
        c = node.Elements.Count
    except Exception:
        try:
            out.append('%s%s = %s' % ('  ' * depth, path, node.Value))
        except Exception:
            out.append('%s%s (leaf)' % ('  ' * depth, path))
        return
    out.append('%s%s  [count=%d]' % ('  ' * depth, path, c))
    for i in range(min(c, limit)):
        try:
            e = node.Elements.Item(i)
        except Exception:
            break
        if e is None:
            out.append('%s  [item %d = None]' % ('  ' * depth, i))
            continue
        try:
            nm = e.Name
        except Exception:
            nm = None
        walk(e, nm if nm else ('item%d' % i), depth + 1, maxdepth, out, limit)
    if c > limit:
        out.append('%s  ... (其余 %d 个省略)' % ('  ' * depth, c - limit))


def probe(path, tag, wait=30):
    print()
    print('=' * 76)
    print('打开:', tag)
    msgs = []

    class Sink:
        def OnControlPanelMessage(self, *a):
            s = ' '.join(str(x) for x in a).strip()
            if s and s != 'False':
                msgs.append(s)

    doc = win32.DispatchEx('Apwn.Document')
    doc.SuppressDialogs = True
    win32.WithEvents(doc, Sink)
    doc.InitFromArchive2(path)
    t = time.time()
    while time.time() - t < wait:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.25)
    print('  消息数:', len(msgs))

    out = []
    for p, md in [(r'\Data\Pure Databanks', 2),
                  (r'\Data\Other Databanks', 2),
                  (r'\Data\Components\UNIFAC-Groups', 3),
                  (r'\Data\Components\Specifications', 2)]:
        out.append('##### ' + p)
        nd = doc.Tree.FindNode(p)
        if nd is None:
            out.append('  None')
        else:
            walk(nd, p.split('\\')[-1], 0, md, out)
        out.append('')

    # NRTL 数据行：找 Input 下所有含 BPVAL/CID 的子节点
    out.append('##### NRTL-1 全子树（深度2）')
    nd = doc.Tree.FindNode(r'\Data\Properties\Parameters\Binary Interaction\NRTL-1')
    if nd is not None:
        walk(nd, 'NRTL-1', 0, 2, out, 40)
    else:
        out.append('  None')

    # 组分规格表内容
    out.append('')
    out.append('##### Components\\Specifications 数据行')
    nd = doc.Tree.FindNode(r'\Data\Components\Specifications\Input')
    if nd is not None:
        try:
            out.append('  count=%d' % nd.Elements.Count)
            for i in range(min(nd.Elements.Count, 20)):
                e = nd.Elements.Item(i)
                out.append('    row: ' + e.Name)
                try:
                    for j in range(min(e.Elements.Count, 8)):
                        ee = e.Elements.Item(j)
                        out.append('        %s = %s' % (ee.Name, ee.Value))
                except Exception:
                    pass
        except Exception as ex:
            out.append('  ' + str(ex)[:80])
    else:
        out.append('  None')

    print('\n'.join(out))

    try:
        doc.Close()
    except Exception:
        pass


probe(os.path.join(BASE, 'NA_work.bkp'), 'NA_work.bkp', 30)
print('DONE')
