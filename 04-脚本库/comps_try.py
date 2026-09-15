# -*- coding: utf-8 -*-
"""顺序试探：找出 NaOH / H2SO4 / Na2SO4 的可解析写法"""
import time
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\x_NA-Chemical-10000t.bkp.backup'
OUT = r'D:\<化工工作区>\_probe\comps_try.bkp'

text = open(SRC, encoding='utf-8', errors='ignore').read()
text = text.replace('RUN-CLASS = PROP', 'RUN-CLASS = FLOWSHEET')
start = text.find('? COMPONENTS MAIN ?')
end = text.find('\\ ? COMPONENTS "ADA/PCS"', start)

ORDER = [
    ('4-MP',  '4-METHYLPYRIDINE'),
    ('NAM',   'NICOTINAMIDE'),
    ('NAC',   'NICOTINIC-ACID'),
    # NaOH 候选
    ('NOHC',  'SODIUM-HYDROXIDE'),
    ('NOHD',  'SODIUM-HYDROXIDE-SOLID'),
    ('NOHE',  'SODIUM OXIDE'),
    ('NOHF',  'CAUSTIC-SODA'),
    # H2SO4 候选
    ('SA1C',  'SULFURIC-ACID'),
    ('SA2C',  'SULFURIC-ACID-1'),
    ('SA3C',  'HYDROGEN-SULFATE'),
    # Na2SO4 候选
    ('SS1C',  'SODIUM-SULFATE'),
    ('SS2C',  'NA2SO4'),
    ('SS3C',  'SODIUM-SULFATE-10'),
    ('SS4C',  'DISODIUM-SULFATE'),
    ('SS5C',  'SODIUM-SULFATE-SOLID'),
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
print('组分 %d 个:' % len(names), [n for n in names if n not in
      ['3-CP', '3-MP', '4-CP', 'AIR', 'CO', 'CO2', 'H2O', 'HCN', 'N2', 'NH3', 'O2', 'TOL']], flush=True)
print()
for c, d in ORDER:
    print('  %-6s %-24s %s' % (c, d, '✓' if c in names else '✗'), flush=True)
try:
    doc.Close()
except Exception:
    pass
