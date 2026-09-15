# -*- coding: utf-8 -*-
"""最终组分尝试（含 CASN 写法）+ 生成修复版文件并验证可加载"""
import time
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\x_NA-Chemical-10000t.bkp.backup'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_修复.bkp'

text = open(SRC, encoding='utf-8', errors='ignore').read()
text = text.replace('RUN-CLASS = PROP', 'RUN-CLASS = FLOWSHEET')
start = text.find('? COMPONENTS MAIN ?')
end = text.find('\\ ? COMPONENTS "ADA/PCS"', start)

ORDER = [
    ('4-MP', '4-METHYLPYRIDINE', None),
    ('NAM',  'NICOTINAMIDE', None),
    ('NAC',  'NICOTINIC-ACID', None),
    ('SS5C', 'SODIUM-SULFATE-SOLID', None),
    ('NOH9', None, '1310-73-2'),
    ('SA09', None, '7664-93-9'),
    ('SS09', None, '7757-82-6'),
    ('NOH1', 'SODIUM-HYDROXIDE', None),
    ('SA01', 'SULFURIC-ACID', None),
]
ents = []
for c, db, cas in ORDER:
    if db:
        ents.append('CID = "%s" OUTNAME = "%s" TYPE = CONV DBNAME1 = "%s"' % (c, c, db))
    else:
        ents.append('CID = "%s" OUTNAME = "%s" TYPE = CONV CASN = "%s"' % (c, c, cas))
text = text[:end] + ' /  ' + ' /  '.join(ents) + ' ' + text[end:]
open(OUT, 'w', encoding='utf-8', errors='ignore').write(text)
print('已写出修复版:', OUT, flush=True)

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
print('组分总数 %d' % len(names), flush=True)
print('新增解析:', [n for n in names if n not in BASE], flush=True)
print()
for c, db, cas in ORDER:
    print('  %-6s %-24s %s' % (c, db or ('CASN ' + cas), '✓' if c in names else '✗'), flush=True)
d = doc.Tree.FindNode(r'\Data')
bl = doc.Tree.FindNode(r'\Data\Blocks')
st = doc.Tree.FindNode(r'\Data\Streams')
print()
print('\\Data 子节点=%s | 模块=%s | 流股=%s' % (
    d.Elements.Count if d else 0, bl.Elements.Count if bl else 0,
    st.Elements.Count if st else 0), flush=True)
try:
    doc.Close()
except Exception:
    pass
