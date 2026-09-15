# -*- coding: utf-8 -*-
import sys, time, os, json
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\final_saved.bkp'
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive2(SRC)
time.sleep(3)


def kids(path):
    n = doc.Tree.FindNode(path)
    if n is None:
        return None
    out = []
    try:
        c = n.Elements.Count
    except Exception:
        return None
    for i in range(c):
        try:
            e = n.Elements.Item(i)
            nm = e.Name
            try:
                v = e.Value
            except Exception:
                v = None
            out.append((nm, v))
        except Exception:
            pass
    return out


print('=== S-106 Output 字段 ===')
kk = kids(r'\Data\Streams\S-106\Output')
if kk:
    for nm, v in kk[:40]:
        print('   %-16s %s' % (nm, v))
else:
    print('   无')

print()
print('=== S-106 Output\\MIXED 字段 ===')
kk = kids(r'\Data\Streams\S-106\Output\MIXED')
if kk:
    for nm, v in kk[:12]:
        print('   %-20s %s' % (nm, v))
else:
    print('   无')

print()
print('=== 试着直接取几个属性 ===')
for p in [r'\Data\Streams\S-106\Output\TEMP_OUT',
          r'\Data\Streams\S-106\Output\PRES_OUT',
          r'\Data\Streams\S-106\Output\MOLEFLOW',
          r'\Data\Streams\S-106\Output\MASSFLMX',
          r'\Data\Streams\S-106\Output\TEMP',
          r'\Data\Streams\S-106\Output\MOLE-FLOW']:
    n = doc.Tree.FindNode(p)
    print('  %-45s -> %s' % (p.split('Output')[-1], 'None' if n is None else (n.Value if True else '')))

try:
    doc.Close()
except Exception:
    pass
print('done')
