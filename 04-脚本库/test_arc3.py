# -*- coding: utf-8 -*-
"""测试 InitFromArchive3 参数 + InitNew2 后 Import"""
import time
import win32com.client as win32

BKP = r'D:\<化工工作区>\_probe\base.bkp'
APW = r'D:\<化工工作区>\_probe\na_main.apw'


def show(doc, tag):
    try:
        d = doc.Tree.FindNode(r'\Data')
        c = d.Elements.Count if d is not None else 0
        s = doc.Tree.FindNode(r'\Data\Streams')
        b = doc.Tree.FindNode(r'\Data\Blocks')
        ms = doc.Tree.FindNode(r'\Data\Properties\Molecular Structure')
        cn = ms.Elements.Count if ms is not None else 0
        print('   %-34s \\Data=%s Streams=%s Blocks=%s 组分=%s' % (
            tag, c, 'None' if s is None else s.Elements.Count,
            'None' if b is None else b.Elements.Count, cn), flush=True)
    except Exception as e:
        print('   %-34s 读取失败 %s' % (tag, str(e)[:60]), flush=True)


# --- A) InitFromArchive3 参数变体 ---
for args in [(BKP,), (BKP, 0), (BKP, 1), (BKP, True), (BKP, 0, 0)]:
    doc = win32.DispatchEx('Apwn.Document')
    try:
        doc.SuppressDialogs = True
    except Exception:
        pass
    try:
        doc.InitFromArchive3(*args)
        time.sleep(2)
        show(doc, 'Archive3(%s)' % ', '.join(['path'] + [repr(a) for a in args[1:]]))
    except Exception as e:
        print('   Archive3 参数%s 失败: %s' % (len(args), str(e)[:70]), flush=True)
    try:
        doc.Close()
    except Exception:
        pass

print(flush=True)
# --- B) InitNew2 + Import ---
for target in [BKP, APW]:
    doc = win32.DispatchEx('Apwn.Document')
    try:
        doc.SuppressDialogs = True
    except Exception:
        pass
    try:
        doc.InitNew2()
        time.sleep(2)
        show(doc, 'InitNew2 (空)')
        doc.Import(target)
        time.sleep(3)
        show(doc, 'InitNew2+Import %s' % target.split('\\')[-1])
    except Exception as e:
        print('   InitNew2+Import 失败: %s' % str(e)[:80], flush=True)
    try:
        doc.Close()
    except Exception:
        pass
