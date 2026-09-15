# -*- coding: utf-8 -*-
"""对照实验：用不同方法打开参考模型/用户模型，检查 Streams、Blocks 是否存在"""
import time
import win32com.client as win32

CASES = [
    ('参考模型', r'D:\<化工工作区>\_probe\ref_multi.bkp'),
    ('用户模型', r'D:\<化工工作区>\_probe\base.bkp'),
]
METHODS = ['InitFromArchive', 'InitFromArchive2']

for tag, path in CASES:
    for m in METHODS:
        print('=' * 20, tag, '|', m)
        doc = win32.DispatchEx('Apwn.Document')
        try:
            getattr(doc, m)(path)
        except Exception as e:
            print('   打开失败:', str(e)[:110])
            continue
        time.sleep(3)
        for probe in [r'\Data', r'\Data\Streams', r'\Data\Blocks',
                      r'\Data\Streams\S-101', r'\Data\Blocks\R-101',
                      r'\Data\Flowsheet']:
            try:
                n = doc.Tree.FindNode(probe)
                if n is None:
                    print('   %-24s None' % probe)
                else:
                    try:
                        c = n.Elements.Count
                    except Exception:
                        c = '-'
                    print('   %-24s 存在  子节点=%s' % (probe, c))
            except Exception as e:
                print('   %-24s ERR %s' % (probe, str(e)[:60]))
        try:
            doc.Close()
        except Exception:
            pass
        print()
