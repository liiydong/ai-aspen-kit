# -*- coding: utf-8 -*-
import sys, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_模拟.bkp'
print('base =', BASE, flush=True)

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive2(BASE)
time.sleep(3)


def kids(path, depth=0, maxd=2):
    n = doc.Tree.FindNode(path)
    if n is None:
        return
    try:
        cnt = n.Elements.Count
    except Exception:
        cnt = 0
    print('%s%s  [%d]' % ('  ' * depth, path.split('\\')[-1], cnt), flush=True)
    if depth >= maxd:
        return
    for i in range(cnt):
        try:
            c = n.Elements.Item(i)
        except Exception:
            continue
        nm = None
        for attr in ('Name', 'Value'):
            try:
                nm = getattr(c, attr)
                break
            except Exception:
                pass
        try:
            cc = c.Elements.Count
        except Exception:
            cc = 0
        print('%s  - %s  [%d]' % ('  ' * depth, nm, cc), flush=True)


print('=== \\Data\\Blocks\\R-101\\Input ===', flush=True)
kids(r'\Data\Blocks\R-101\Input', 0, 1)

for p in [r'\Data\Blocks\R-101\Input\STOIC',
          r'\Data\Blocks\R-101\Input\COEF',
          r'\Data\Blocks\R-101\Input\CONVEX',
          r'\Data\Blocks\R-101\Input\CONV']:
    n = doc.Tree.FindNode(p)
    print('---', p, '->', 'None' if n is None else 'OK', flush=True)
    if n is not None:
        try:
            print('     count =', n.Elements.Count, flush=True)
            for i in range(min(n.Elements.Count, 12)):
                e = n.Elements.Item(i)
                print('       [%d] %s = %s' % (i, e.Name, e.Value), flush=True)
        except Exception as ex:
            print('     err', ex, flush=True)

try:
    doc.Close()
except Exception:
    pass
print('done', flush=True)
