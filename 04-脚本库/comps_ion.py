# -*- coding: utf-8 -*-
"""去掉 DATABANKS 改动；dump 数据库节点；试离子型与溶液型组分名"""
import time
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\x_NA-Chemical-10000t.bkp.backup'
OUT = r'D:\<化工工作区>\_probe\comps_ion.bkp'

text = open(SRC, encoding='utf-8', errors='ignore').read()
text = text.replace('RUN-CLASS = PROP', 'RUN-CLASS = FLOWSHEET')
start = text.find('? COMPONENTS MAIN ?')
end = text.find('\\ ? COMPONENTS "ADA/PCS"', start)

ORDER = [
    ('4-MP', '4-METHYLPYRIDINE'),
    ('NAM',  'NICOTINAMIDE'),
    ('NAC',  'NICOTINIC-ACID'),
    ('SS01', 'SODIUM-SULFATE-SOLID'),
    ('NA+',  'NA+'),
    ('OH-',  'OH-'),
    ('H3O+', 'H3O+'),
    ('SO4--', 'SO4--'),
    ('HSO4-', 'HSO4-'),
    ('NOH7', 'SODIUM-HYDROXIDE-SOLUTION'),
    ('NOH8', 'NAOH-SOLUTION'),
    ('SA07', 'SULFURIC-ACID-SOLUTION'),
]
entries = ['CID = "%s" OUTNAME = "%s" TYPE = CONV DBNAME1 = "%s"' % (c, c, d) for c, d in ORDER]
text = text[:end] + ' /  ' + ' /  '.join(entries) + ' ' + text[end:]
open(OUT, 'w', encoding='utf-8', errors='ignore').write(text)

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(OUT)
time.sleep(4)

ms = doc.Tree.FindNode(r'\Data\Properties\Molecular Structure')
names = [ms.Elements.Item(i).Name for i in range(ms.Elements.Count)] if ms else []
BASE = ['3-CP', '3-MP', '4-CP', 'AIR', 'CO', 'CO2', 'H2O', 'HCN', 'N2', 'NH3', 'O2', 'TOL']
print('新增:', [n for n in names if n not in BASE], flush=True)
print()
for c, d in ORDER:
    print('  %-6s %-28s %s' % (c, d, '✓' if c in names else '✗'), flush=True)

print()
print('=== 数据库节点 ===', flush=True)
for path in [r'\Data\Pure Databanks\Input', r'\Data\Other Databanks\Input']:
    n = doc.Tree.FindNode(path)
    print('[%s] %s' % (path, 'None' if n is None else n.Elements.Count), flush=True)
    if n is None:
        continue
    for i in range(n.Elements.Count):
        ch = n.Elements.Item(i)
        try:
            print('   %-22s = %r' % (ch.Name, ch.Value), flush=True)
        except Exception as e:
            print('   %-22s 读取失败(%s)' % (ch.Name, str(e)[:40]), flush=True)
try:
    doc.Close()
except Exception:
    pass
