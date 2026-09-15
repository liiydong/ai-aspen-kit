# -*- coding: utf-8 -*-
"""按序注入组分（遇到坏条目会中断，所以顺序要对）"""
import time, shutil
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\x_NA-Chemical-10000t.bkp.backup'
OUT = r'D:\<化工工作区>\_probe\comps_final.bkp'

text = open(SRC, encoding='utf-8', errors='ignore').read()
text = text.replace('RUN-CLASS = PROP', 'RUN-CLASS = FLOWSHEET')

start = text.find('? COMPONENTS MAIN ?')
end = text.find('\\ ? COMPONENTS "ADA/PCS"', start)

ORDER = [
    ('4-MP',   '4-METHYLPYRIDINE'),
    ('NAM',    'NICOTINAMIDE'),
    ('NAC',    'NICOTINIC-ACID'),
    ('NAOH',   'SODIUM-HYDROXIDE'),
    ('H2SO4',  'SULFURIC-ACID'),
    ('SS1',    'SODIUM-SULFATE'),
    ('SS2',    'NA2SO4'),
    ('SS3',    'DISODIUM-SULFATE'),
    ('SS4',    'SODIUM-SULFATE-10'),
    ('SS5',    'SODIUM-SULPHATE'),
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
print('最终组分 %d 个:' % len(names), names, flush=True)
print()
for c, d in ORDER:
    print('  %-6s %-22s %s' % (c, d, '✓' if c in names else '✗'), flush=True)
st = doc.Tree.FindNode(r'\Data\Streams')
bl = doc.Tree.FindNode(r'\Data\Blocks')
print()
print('流股:', 'None' if st is None else st.Elements.Count,
      '| 模块:', 'None' if bl is None else bl.Elements.Count, flush=True)
try:
    doc.Close()
except Exception:
    pass
