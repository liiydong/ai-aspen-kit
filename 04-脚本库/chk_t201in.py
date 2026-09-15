# -*- coding: utf-8 -*-
import sys, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive2(r'D:\<化工工作区>\NA-Chemical-10000t_sim6.bkp')
time.sleep(3)


def walk(path, maxd=2, d=0):
    n = doc.Tree.FindNode(path)
    if n is None:
        print('%s%s -> None' % ('  ' * d, path.split('\\')[-1])); return
    val = None
    try:
        val = n.Value
    except Exception:
        pass
    try:
        cnt = n.Elements.Count
    except Exception:
        cnt = -1
    print('%s%s [%d] val=%r' % ('  ' * d, path.split('\\')[-1], cnt, val))
    if d >= maxd or cnt <= 0:
        return
    for i in range(cnt):
        try:
            e = n.Elements.Item(i)
        except Exception:
            continue
        nm = None
        try:
            nm = e.Name
        except Exception:
            pass
        v = None
        try:
            v = e.Value
        except Exception:
            pass
        try:
            c = e.Elements.Count
        except Exception:
            c = -1
        print('%s  .%s [%d] val=%r' % ('  ' * d, nm, c, v))
        if c > 0 and d + 1 <= maxd:
            for j in range(c):
                try:
                    e2 = e.Elements.Item(j)
                except Exception:
                    continue
                nm2 = None
                v2 = None
                try:
                    nm2 = e2.Name
                except Exception:
                    pass
                try:
                    v2 = e2.Value
                except Exception:
                    pass
                print('%s    ..%s val=%r' % ('  ' * d, nm2, v2))


for p in [r'\Data\Blocks\T-201\Input\FEEDS',
          r'\Data\Blocks\T-201\Input\PRODUCTS',
          r'\Data\Blocks\T-201\Input\NSTAGE',
          r'\Data\Blocks\T-201\Input\PRES1',
          r'\Data\Blocks\T-201\Input\RDV',
          r'\Data\Blocks\T-201\Input\BASIS_RDV',
          r'\Data\Blocks\T-201\Input\RATE_RDV',
          r'\Data\Blocks\T-201\Input\QN',
          r'\Data\Blocks\T-201\Input\CONDENSER',
          r'\Data\Blocks\T-201\Input\REBOILER']:
    print('#' * 12, p)
    walk(p)
    print()

try:
    doc.Close()
except Exception:
    pass
print('done')
