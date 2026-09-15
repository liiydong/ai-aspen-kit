# -*- coding: utf-8 -*-
"""扫描所有可能的加载方法，找能读出 Streams/Blocks 的那个"""
import time
import win32com.client as win32

TARGET = r'D:\<化工工作区>\_probe\na_main.apw'
METHODS = ['InitFromFile2', 'InitFromXML', 'InitFromTemplate2', 'Restore',
           'Restore2', 'Readback', 'LoadLink', 'New2', 'New3', 'InitNew2',
           'InitFromArchive', 'InitFromArchive2']

for m in METHODS:
    doc = win32.DispatchEx('Apwn.Document')
    try:
        doc.SuppressDialogs = True
    except Exception:
        pass
    fn = getattr(doc, m, None)
    if fn is None:
        print('%-20s 不存在' % m, flush=True)
        continue
    try:
        fn(TARGET)
    except Exception as e:
        print('%-20s 调用失败 %s' % (m, str(e)[:70]), flush=True)
        try:
            doc.Close()
        except Exception:
            pass
        continue
    time.sleep(2)
    try:
        d = doc.Tree.FindNode(r'\Data')
        c = d.Elements.Count if d is not None else 0
        s = doc.Tree.FindNode(r'\Data\Streams')
        b = doc.Tree.FindNode(r'\Data\Blocks')
        msg = '\\Data=%s Streams=%s Blocks=%s' % (
            c, 'None' if s is None else s.Elements.Count,
            'None' if b is None else b.Elements.Count)
        print('%-20s OK   %s' % (m, msg), flush=True)
    except Exception as e:
        print('%-20s 树读取失败 %s' % (m, str(e)[:70]), flush=True)
    try:
        doc.Close()
    except Exception:
        pass
