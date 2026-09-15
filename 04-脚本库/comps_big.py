# -*- coding: utf-8 -*-
"""综合测试：DATABANKS 清单 + 大批候选组分名"""
import time
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\x_NA-Chemical-10000t.bkp.backup'
OUT = r'D:\<化工工作区>\_probe\comps_big.bkp'

text = open(SRC, encoding='utf-8', errors='ignore').read()
text = text.replace('RUN-CLASS = PROP', 'RUN-CLASS = FLOWSHEET')

# --- DATABANKS ---
i = text.find('? DATABANKS ?')
j = text.find('\\ DATABANKS', i)
m = text.find('? COMPONENTS MAIN ?', i)
DSET = ('\nFILE-SYM-NAM = ( "APV150 PURE41" "APV150 AQUEOUS" "APV150 SOLIDS" '
        '"APV150 INORGANIC" "APESV150 AP-EOS" "NISTV150 NIST-TRC" ) \\ ')
text = text[:j + len('\\ DATABANKS')] + DSET + text[m:]
print('DATABANKS 已补', flush=True)

# --- 组分 ---
start = text.find('? COMPONENTS MAIN ?')
end = text.find('\\ ? COMPONENTS "ADA/PCS"', start)

ORDER = [
    ('4-MP', '4-METHYLPYRIDINE'),
    ('NAM',  'NICOTINAMIDE'),
    ('NAC',  'NICOTINIC-ACID'),
    # NaOH
    ('NOH1', 'SODIUM-HYDROXIDE'),
    ('NOH2', 'SODIUM-HYDROXIDE-SOLID'),
    ('NOH3', 'SODIUM-HYDROXIDE-AQ'),
    ('NOH4', 'SODIUM-HYDROXIDE-1'),
    ('NOH5', 'NAOH-SOLID'),
    ('NOH6', 'SODIUMHYDROXIDE'),
    # H2SO4
    ('SA01', 'SULFURIC-ACID'),
    ('SA02', 'SULFURIC-ACID-SOLID'),
    ('SA03', 'SULFURIC-ACID-AQ'),
    ('SA04', 'SULFURIC-ACID-1'),
    ('SA05', 'H2SO4-SOLID'),
    ('SA06', 'SULFURICACID'),
    # Na2SO4（已确认）
    ('SS01', 'SODIUM-SULFATE-SOLID'),
    ('SS02', 'SODIUM-SULFATE'),
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
print('新增解析出的:', [n for n in names if n not in BASE], flush=True)
print()
for c, d in ORDER:
    print('  %-6s %-26s %s' % (c, d, '✓' if c in names else '✗'), flush=True)
print()
bl = doc.Tree.FindNode(r'\Data\Blocks')
st = doc.Tree.FindNode(r'\Data\Streams')
print('模块 %s | 流股 %s' % (bl.Elements.Count if bl else 0, st.Elements.Count if st else 0), flush=True)
try:
    doc.Close()
except Exception:
    pass
