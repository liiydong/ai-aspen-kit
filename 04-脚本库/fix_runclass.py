# -*- coding: utf-8 -*-
"""修复 RUN-CLASS = PROP -> FLOWSHEET，并补齐组分，验证流程图是否加载"""
import time
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\x_NA-Chemical-10000t.bkp.backup'
OUT = r'D:\<化工工作区>\_probe\fixed_runclass.bkp'

text = open(SRC, encoding='utf-8', errors='ignore').read()

# --- 1) 修 RUN-CLASS ---
before = text.count('RUN-CLASS = PROP')
text = text.replace('RUN-CLASS = PROP', 'RUN-CLASS = FLOWSHEET')
# 补 MOLEFLOW（参考文件有，用于流股流量基准）
if 'MOLEFLOW = MOLEFLOW' not in text:
    text = text.replace('SIMULATE INTERACTIVE = NO TFFFILE',
                        'SIMULATE INTERACTIVE = NO MOLEFLOW = MOLEFLOW TFFFILE')
print('RUN-CLASS = PROP 出现次数:', before)
print('MOLEFLOW 已补:', 'MOLEFLOW = MOLEFLOW' in text, flush=True)

# --- 2) 补组分 ---
start = text.find('? COMPONENTS MAIN ?')
end = text.find('\\ ? COMPONENTS "ADA/PCS"', start)
NEW = [
    ('4-MP',  'DBNAME1 = "4-METHYLPYRIDINE"'),
    ('NAM',   'DBNAME1 = "NICOTINAMIDE"'),
    ('NAC',   'DBNAME1 = "NICOTINIC-ACID"'),
    ('NAOH',  'DBNAME1 = "SODIUM-HYDROXIDE"'),
    ('H2SO4', 'DBNAME1 = "SULFURIC-ACID"'),
]
entries = ['CID = "%s" OUTNAME = "%s" TYPE = CONV %s' % (c, c, d) for c, d in NEW]
text = text[:end] + ' /  ' + ' /  '.join(entries) + ' ' + text[end:]
open(OUT, 'w', encoding='utf-8', errors='ignore').write(text)
print('已写出', OUT, flush=True)

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(OUT)
time.sleep(4)

d = doc.Tree.FindNode(r'\Data')
print()
print('>> \\Data 子节点数:', d.Elements.Count if d else 0, flush=True)
if d:
    print('>> 子节点:', [d.Elements.Item(i).Name for i in range(d.Elements.Count)], flush=True)
for p in [r'\Data\Streams', r'\Data\Blocks']:
    n = doc.Tree.FindNode(p)
    if n is None:
        print('>> %-16s None' % p, flush=True)
    else:
        print('>> %-16s 存在 %s 个: %s' % (
            p, n.Elements.Count,
            [n.Elements.Item(i).Name for i in range(min(20, n.Elements.Count))]), flush=True)
ms = doc.Tree.FindNode(r'\Data\Properties\Molecular Structure')
if ms:
    print('>> 组分 %d: %s' % (ms.Elements.Count,
          [ms.Elements.Item(i).Name for i in range(ms.Elements.Count)]), flush=True)
try:
    doc.Close()
except Exception:
    pass
