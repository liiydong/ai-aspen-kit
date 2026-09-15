# -*- coding: utf-8 -*-
"""在正确位置（COMPONENTS MAIN 之后）插入组分"""
import time, sys
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\na_text.bkp'
DST = r'D:\<化工工作区>\_probe\test_comp2.bkp'

comp = [
    ('3-MP',   '3-METHYLPYRIDINE',      '108-99-6'),
    ('4-MP',   '4-METHYLPYRIDINE',      '108-89-4'),
    ('NH3',    'AMMONIA',               '7664-41-7'),
    ('O2',     'OXYGEN',                '7782-44-7'),
    ('N2',     'NITROGEN',              '7727-37-9'),
    ('WATER',  'WATER',                 '7732-18-5'),
    ('3-CP',   '3-CYANOPYRIDINE',       '100-54-9'),
    ('4-CP',   '4-CYANOPYRIDINE',       '100-48-1'),
    ('TOLUENE','TOLUENE',               '108-88-3'),
    ('NAM',    'NICOTINAMIDE',          '98-92-0'),
    ('NA',     'NICOTINIC-ACID',        '59-67-6'),
    ('NAOH',   'SODIUM-HYDROXIDE',      '1310-73-2'),
    ('H2SO4',  'SULFURIC-ACID',         '7664-93-9'),
    ('NA2SO4', 'SODIUM-SULFATE',        '7757-82-6'),
    ('CO2',    'CARBON-DIOXIDE',        '124-38-9'),
    ('CO',     'CARBON-MONOXIDE',       '630-08-0'),
    ('HCN',    'HYDROGEN-CYANIDE',      '74-90-8'),
]

def entry(cid, dbname, alias):
    return (f'CID = "{cid}" ANAME = "{alias}" OUTNAME = "{cid}" TYPE = CONV '
            f'DBNAME1 = "{dbname}" ANAME1 = "{alias}" /  ')

body = '\\ COMPONENTS ' + ''.join(entry(*c) for c in comp)

text = open(SRC, encoding='utf-8', errors='ignore').read()
anchor = '? COMPONENTS MAIN ?'
idx = text.find(anchor)
print('锚点位置:', idx)
# 显示锚点附近
print('锚点附近原文:', repr(text[idx:idx + 110]))
# 在锚点之后紧跟一个空格再插入
new = text[:idx + len(anchor)] + ' ' + body + text[idx + len(anchor):]
open(DST, 'w', encoding='utf-8').write(new)
print('已写:', DST)

doc = win32.DispatchEx('Apwn.Document')
ok = False
for m in ('InitFromArchive2', 'InitFromArchive'):
    try:
        getattr(doc, m)(DST); print(f'{m} 成功'); ok = True; break
    except Exception as e:
        print(f'{m} 失败: {str(e)[:130]}')
if not ok:
    sys.exit(1)

time.sleep(4)
c = doc.Tree.FindNode(r'\Data\Components\Specifications\Input')
print('组分节点元素数:', c.Elements.Count)
print('子节点:', [c.Elements.Item(i).Name for i in range(min(20, c.Elements.Count))])

try:
    p = r'\Data\Streams\S-101\Input\TEMP'
    doc.Tree.FindNode(p).Value = 25.0
    print('>>> 流股温度写入成功:', doc.Tree.FindNode(p).Value)
except Exception as e:
    print('>>> 流股温度写入失败:', str(e)[:140])
