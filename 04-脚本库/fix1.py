# -*- coding: utf-8 -*-
"""补 DATABANKS + 补齐组分，验证流程图是否随之加载"""
import time, sys
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\x_NA-Chemical-10000t.bkp.backup'
OUT = r'D:\<化工工作区>\_probe\fix1.bkp'

text = open(SRC, encoding='utf-8', errors='ignore').read()

# ---- 1) 补 DATABANKS ----
i = text.find('? DATABANKS ?')
j = text.find('\\ DATABANKS', i)
m = text.find('? COMPONENTS MAIN ?', i)
print('DATABANKS 段: i=%d j=%d m=%d' % (i, j, m))
print('  被替换的片段:', repr(text[j + len('\\ DATABANKS'):m]))
DSET = ('\nFILE-SYM-NAM = ( "APV150 PURE41" "APV150 AQUEOUS" "APV150 SOLIDS" '
        '"APV150 INORGANIC" "APESV150 AP-EOS" "NISTV150 NIST-TRC" ) \\ ')
text = text[:j + len('\\ DATABANKS')] + DSET + text[m:]

# ---- 2) 补组分 ----
start = text.find('? COMPONENTS MAIN ?')
marker = '\\ ? COMPONENTS "ADA/PCS"'
end = text.find(marker, start)
NEW = [
    ('4-MP',   'DBNAME1 = "4-METHYLPYRIDINE"'),
    ('NAM',    'DBNAME1 = "NICOTINAMIDE"'),
    ('NAC',    'DBNAME1 = "NICOTINIC-ACID"'),
    ('NAOH',   'DBNAME1 = "SODIUM-HYDROXIDE"'),
    ('H2SO4',  'DBNAME1 = "SULFURIC-ACID"'),
    ('NAS1',   'DBNAME1 = "SODIUM-SULFATE"'),
    ('NAS2',   'DBNAME1 = "NA2SO4"'),
    ('NAS3',   'DBNAME1 = "DISODIUM-SULFATE"'),
    ('NAS4',   'DBNAME1 = "SODIUM-SULFATE-10"'),
]
entries = ['CID = "%s" OUTNAME = "%s" TYPE = CONV %s' % (c, c, d) for c, d in NEW]
text = text[:end] + ' /  ' + ' /  '.join(entries) + ' ' + text[end:]
open(OUT, 'w', encoding='utf-8', errors='ignore').write(text)
print('已写出', OUT, len(text), flush=True)

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(OUT)
time.sleep(3)
ms = doc.Tree.FindNode(r'\Data\Properties\Molecular Structure')
names = [ms.Elements.Item(i).Name for i in range(ms.Elements.Count)] if ms else []
print('组分 %d: %s' % (len(names), names), flush=True)
for p in [r'\Data', r'\Data\Streams', r'\Data\Blocks', r'\Data\Flowsheet']:
    n = doc.Tree.FindNode(p)
    print('  %-16s %s' % (p, 'None' if n is None else '存在(子=%s)' % n.Elements.Count), flush=True)
try:
    doc.Close()
except Exception:
    pass
