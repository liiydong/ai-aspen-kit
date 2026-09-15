# -*- coding: utf-8 -*-
import sys, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive2(r'D:\<化工工作区>\_probe\final_saved.bkp')
time.sleep(3)

print('=== 各流股汇总字段 ===')
for s in ['S-104', 'S-106', 'S-108', 'S-110', 'S-112', 'S-114', 'S-116', 'S-121',
          'S-122', 'S-201', 'S-301', 'S-402']:
    b = r'\Data\Streams\%s\Output' % s
    row = []
    for f in ['TEMP_OUT', 'PRES_OUT', 'MOLEFLOW', 'MASSFLMX', 'BETA', 'VFRAC', 'MASSVFRA']:
        n = doc.Tree.FindNode(b + '\\' + f)
        row.append('%s=%s' % (f, 'None' if n is None else n.Value))
    print('  %-7s %s' % (s, ' | '.join(row)))

print()
print('=== 找组分流量容器 ===')
for p in [r'\Data\Streams\S-106\Output\MOLEFLOW',
          r'\Data\Streams\S-106\Output\MOLE-FLOW',
          r'\Data\Streams\S-106\Output\MASSFLOW',
          r'\Data\Streams\S-106\Output\HMX_FLOW',
          r'\Data\Streams\S-106\Output\MIXED']:
    n = doc.Tree.FindNode(p)
    if n is None:
        print('  %-45s None' % p.split('Output')[-1]); continue
    try:
        c = n.Elements.Count
    except Exception:
        c = -1
    print('  %-45s [%d]' % (p.split('Output')[-1], c))
    for i in range(min(c, 20)):
        try:
            e = n.Elements.Item(i)
            print('        .%s = %s' % (e.Name, e.Value))
        except Exception:
            pass

print()
print('=== T-401 输出 ===')
b = r'\Data\Blocks\T-401\Output'
for f in ['TEMP', 'PRES', 'COND_DUTY', 'REB_DUTY', 'MOLE_RR', 'BOTTOM_TEMP', 'TOP_TEMP']:
    n = doc.Tree.FindNode(b + '\\' + f)
    print('  %-12s %s' % (f, 'None' if n is None else n.Value))

try:
    doc.Close()
except Exception:
    pass
print('done')
