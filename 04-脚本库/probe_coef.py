# -*- coding: utf-8 -*-
import sys, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_模拟.bkp'

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive2(BASE)
time.sleep(3)


def walk(path, maxd=3, depth=0):
    n = doc.Tree.FindNode(path)
    if n is None:
        print('%s%s -> None' % ('  ' * depth, path.split('\\')[-1]))
        return
    try:
        cnt = n.Elements.Count
    except Exception:
        cnt = -1
    val = None
    try:
        val = n.Value
    except Exception:
        pass
    print('%s%s  [%d]  val=%r' % ('  ' * depth, path.split('\\')[-1], cnt, val))
    if depth >= maxd:
        return
    for i in range(max(0, cnt)):
        try:
            e = n.Elements.Item(i)
        except Exception:
            continue
        nm = None
        for a in ('Name',):
            try:
                nm = getattr(e, a)
            except Exception:
                pass
        v2 = None
        try:
            v2 = e.Value
        except Exception:
            pass
        try:
            c2 = e.Elements.Count
        except Exception:
            c2 = -1
        print('%s  .%s  [%d] val=%r' % ('  ' * depth, nm, c2, v2))
        if c2 > 0 and depth + 1 <= maxd:
            for j in range(min(c2, 10)):
                try:
                    e2 = e.Elements.Item(j)
                except Exception:
                    continue
                nm2 = None
                try:
                    nm2 = e2.Name
                except Exception:
                    pass
                v3 = None
                try:
                    v3 = e2.Value
                except Exception:
                    pass
                try:
                    c3 = e2.Elements.Count
                except Exception:
                    c3 = -1
                print('%s    ..%s [%d] val=%r' % ('  ' * depth, nm2, c3, v3))


for f in ['COEF', 'COEF1', 'CONV', 'EXTENT', 'KEY_CID', 'KEY_SSID', 'PROD_PHASE', 'OPT_EXT_CONV']:
    print('#' * 20, f)
    walk(r'\Data\Blocks\R-101\Input\%s' % f, maxd=3)
    print()

try:
    doc.Close()
except Exception:
    pass
print('done')
