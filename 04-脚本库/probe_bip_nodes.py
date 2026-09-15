# -*- coding: utf-8 -*-
"""用正确树路径：读内存中的二元参数条数 + UNIFAC 基团 + DATABANKS 节点"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>'


def show(doc, path, limit=30):
    nd = doc.Tree.FindNode(path)
    if nd is None:
        print('  %-62s None' % path)
        return None
    try:
        c = nd.Elements.Count
    except Exception as e:
        try:
            print('  %-62s Value=%s' % (path, nd.Value))
        except Exception:
            print('  %-62s (no Elements)' % path)
        return nd
    print('  %-62s count=%d' % (path, c))
    for i in range(min(c, limit)):
        try:
            print('        -', nd.Elements.Item(i).Name)
        except Exception:
            pass
    return nd


def probe(path, tag, wait=25):
    print()
    print('=' * 74)
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
    # 等输入翻译
    t = time.time()
    while time.time() - t < wait:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.25)

    print('  等待 %.0fs，消息数=%d' % (wait, len(msgs)))
    for m in msgs:
        if re.search(r'ERROR|SEVERE|CANNOT|MISSING', m, re.I):
            print('     !', m[:130])

    print('  --- 顶层 Data ---')
    show(doc, r'\Data', 30)

    print('  --- 二元参数 ---')
    for p in [r'\Data\Properties\Parameters\Binary Interaction',
              r'\Data\Properties\Parameters\Binary Interaction\NRTL-1',
              r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Input']:
        show(doc, p, 20)

    print('  --- UNIFAC 基团 ---')
    for p in [r'\Data\Properties\Parameters\UNIFAC Groups',
              r'\Data\Properties\Parameters\UNIFAC Groups\Input',
              r'\Data\Properties\Parameters\UNIFAC Groups Binary']:
        show(doc, p, 20)

    # 组分
    print('  --- 组分 ---')
    show(doc, r'\Data\Components', 25)

    # 读 NRTL 参数条数（Input 的 Elements）
    nd = doc.Tree.FindNode(r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Input')
    if nd is not None:
        try:
            print('  NRTL-1\\Input 条数:', nd.Elements.Count)
            for i in range(min(nd.Elements.Count, 20)):
                e = nd.Elements.Item(i)
                print('      row:', e.Name)
                try:
                    for j in range(min(e.Elements.Count, 15)):
                        ee = e.Elements.Item(j)
                        print('          %s = %s' % (ee.Name, ee.Value))
                except Exception:
                    pass
        except Exception as e2:
            print('  ', str(e2)[:100])

    try:
        doc.Close()
    except Exception:
        pass


probe(os.path.join(BASE, 'NA_work.bkp'), 'NA_work.bkp（基线，DATABANKS 空）', 30)
print('DONE')
