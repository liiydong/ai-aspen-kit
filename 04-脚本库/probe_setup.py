# -*- coding: utf-8 -*-
import sys, time, re
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_BIP2.bkp'
doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
doc.InitFromArchive2(SRC)
time.sleep(2)


def dump(path, depth, maxd=2):
    try:
        n = doc.Tree.FindNode(path)
    except Exception as e:
        print('  %s ERR %s' % (path, str(e)[:40]))
        return
    if n is None:
        print('  %s -> None' % path)
        return
    try:
        c = n.Elements.Count
    except Exception:
        try:
            print('  %s = %s' % (path, str(n.Value)[:40]))
        except Exception:
            pass
        return
    for i in range(c):
        try:
            e = n.Elements.Item(i)
        except Exception:
            continue
        nm = e.Name
        try:
            v = e.Value
        except Exception:
            v = ''
        try:
            cc = e.Elements.Count
        except Exception:
            cc = 0
        print('%s%s  (n=%s) val=%s' % ('  ' * depth, nm, cc, str(v)[:30]))
        if cc and depth < maxd:
            dump(path + '\\' + nm, depth + 1, maxd)


print('=== \\Data\\Setup 两层 ===')
dump(r'\Data\Setup', 0, 2)
print()
print('=== \\Data\\Components 一层 ===')
dump(r'\Data\Components', 0, 0)
print()
print('=== \\Data\\Components\\Specifications\\Input 里含 CID 的行数 ===')
try:
    n = doc.Tree.FindNode(r'\Data\Components\Specifications\Input\CID')
    print('  CID rows =', n.Elements.Count if n is not None else None)
except Exception as e:
    print('  ERR', str(e)[:60])
try:
    doc.Close()
except Exception:
    pass
print('DONE')
