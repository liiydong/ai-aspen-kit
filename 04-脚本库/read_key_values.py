# -*- coding: utf-8 -*-
"""读取关键值：FILE_SYM_NAM / BDBANK / ESTIMATE / ACCESSDB / LOADDECH / 分子式"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>'


def dump(node, path, indent, out, maxn=30):
    if node is None:
        out.append('%s%s = None' % (indent, path))
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
            out.append('%s%s = %r  (空)' % (indent, path, node.Value))
        except Exception:
            out.append('%s%s = (空表)' % (indent, path))
        return
    out.append('%s%s  [%d]' % (indent, path, c))
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
        dump(e, nm, indent + '   ', out, maxn)


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
    for p in [r'\Data\Pure Databanks\Input\FILE_SYM_NAM',
              r'\Data\Other Databanks\Input\FILE_SYM_NAM',
              r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Input\BDBANK',
              r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Input\ESTIMATE',
              r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Input\NEL',
              r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\CC Nodes',
              r'\Data\Components\Specifications\Output']:
        node = doc.Tree.FindNode(p)
        dump(node, p.split('\\')[-1], '', out, 25)
        out.append('')

    print('\n'.join(out))
    try:
        doc.Close()
    except Exception:
        pass


probe(os.path.join(BASE, 'NA_work.bkp'), 'NA_work.bkp', 30)
print('DONE')
