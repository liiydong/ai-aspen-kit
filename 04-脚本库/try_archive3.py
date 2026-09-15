# -*- coding: utf-8 -*-
"""测试 InitFromArchive3 / InitFromXML 等，并检查组分与流程图是否加载"""
import time
import win32com.client as win32

TARGETS = [
    ('apwz', r'D:\<化工工作区>\NA-Chemical-10000t.apwz'),
    ('apw',  r'D:\<化工工作区>\_probe\na_main.apw'),
    ('bkp',  r'D:\<化工工作区>\_probe\base.bkp'),
]

for tag, path in TARGETS:
    for m in ['InitFromArchive3', 'InitFromArchive2']:
        print('=' * 20, tag, '|', m)
        doc = win32.DispatchEx('Apwn.Document')
        try:
            doc.SuppressDialogs = True
        except Exception:
            pass
        try:
            getattr(doc, m)(path)
        except Exception as e:
            print('   调用失败:', str(e)[:150])
            continue
        time.sleep(3)
        for p in [r'\Data', r'\Data\Streams', r'\Data\Blocks', r'\Data\Flowsheet']:
            n = doc.Tree.FindNode(p)
            if n is None:
                print('   %-18s None' % p)
            else:
                try:
                    c = n.Elements.Count
                except Exception:
                    c = '-'
                print('   %-18s 存在 子=%s' % (p, c))
        # 组分个数
        try:
            ms = doc.Tree.FindNode(r'\Data\Properties\Molecular Structure')
            print('   组分数:', ms.Elements.Count if ms else 'None')
        except Exception as e:
            print('   组分读取失败:', str(e)[:80])
        try:
            doc.Close()
        except Exception:
            pass
        print()
