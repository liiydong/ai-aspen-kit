# -*- coding: utf-8 -*-
"""尝试用各种方法打开二进制 .apw"""
import time
import win32com.client as win32

APW = r'D:\<化工工作区>\_probe\na_main.apw'

doc = win32.DispatchEx('Apwn.Document')
ms = [m for m in dir(doc) if not m.startswith('_')]
print('Apwn.Document 方法/属性:')
print('  ', ms)
print()

for m in ['InitFromArchive', 'InitFromArchive2', 'Open', 'InitFromFile']:
    if not hasattr(doc, m):
        print('%-18s 不存在' % m)
        continue
    d2 = win32.DispatchEx('Apwn.Document')
    try:
        getattr(d2, m)(APW)
        print('%-18s 调用成功' % m)
        time.sleep(3)
        for p in [r'\Data', r'\Data\Streams', r'\Data\Blocks']:
            n = d2.Tree.FindNode(p)
            if n is None:
                print('     %-16s None' % p)
            else:
                try:
                    c = n.Elements.Count
                except Exception:
                    c = '-'
                print('     %-16s 存在 子=%s' % (p, c))
        try:
            d2.Close()
        except Exception:
            pass
    except Exception as e:
        print('%-18s 失败: %s' % (m, str(e)[:140]))
    print()
