# -*- coding: utf-8 -*-
"""枚举法解析剩余组分：NaOH / H2SO4 / Na2SO4"""
import time
import win32com.client as win32

BASE = r'D:\<化工工作区>\_probe\fixed_runclass.bkp'   # 已修 RUN-CLASS + 含 4-MP/NAM/NAC
OUT = r'D:\<化工工作区>\_probe\comp_sweep.bkp'

text = open(BASE, encoding='utf-8', errors='ignore').read()

# 先看当前组分段是否规范
start = text.find('? COMPONENTS MAIN ?')
end = text.find('\\ ? COMPONENTS "ADA/PCS"', start)
print('=== 组分段尾部 300 字符 ===')
print(repr(text[end - 300:end]))
print()

CANDS = [
    ('NOH1', 'SODIUM-HYDROXIDE'),
    ('NOH2', 'NAOH'),
    ('SA1',  'SULFURIC-ACID'),
    ('SA2',  'H2SO4'),
    ('SS1',  'SODIUM-SULFATE'),
    ('SS2',  'NA2SO4'),
    ('SS3',  'DISODIUM-SULFATE'),
    ('SS4',  'SODIUM-SULFATE-10'),
]
entries = ['CID = "%s" OUTNAME = "%s" TYPE = CONV DBNAME1 = "%s"' % (c, c, d) for c, d in CANDS]
text = text[:end] + ' /  ' + ' /  '.join(entries) + ' ' + text[end:]
open(OUT, 'w', encoding='utf-8', errors='ignore').write(text)
print('待测:', [c for c, _ in CANDS], flush=True)

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(OUT)
time.sleep(4)
ms = doc.Tree.FindNode(r'\Data\Properties\Molecular Structure')
names = [ms.Elements.Item(i).Name for i in range(ms.Elements.Count)] if ms else []
print('组分 %d: %s' % (len(names), names), flush=True)
print()
for c, d in CANDS:
    print('  %-6s (%s): %s' % (c, d, '✓ 解析成功' if c in names else '✗ 失败'), flush=True)
try:
    doc.Close()
except Exception:
    pass
