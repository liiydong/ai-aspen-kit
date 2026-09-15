# -*- coding: utf-8 -*-
"""在文本 bkp 中插入 COMPONENTS 段，然后用 Aspen 验证"""
import re, time, sys
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\na_text.bkp'
DST = r'D:\<化工工作区>\_probe\test_comp.bkp'

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

body = '\\ ? COMPONENTS MAIN ? \\ COMPONENTS ' + ''.join(entry(*c) for c in comp)
body = body.rstrip() + '\n'

text = open(SRC, encoding='utf-8', errors='ignore').read()
anchor = '? SETUP MAIN ?'
if anchor not in text:
    print('找不到 SETUP 锚点，退出'); sys.exit(1)
text2 = text.replace(anchor, body + anchor, 1)
open(DST, 'w', encoding='utf-8').write(text2)
print('已写入:', DST, len(text2), 'chars')

doc = win32.DispatchEx('Apwn.Document')
for m in ('InitFromArchive2', 'InitFromArchive'):
    try:
        getattr(doc, m)(DST)
        print(f'{m} 打开成功')
        break
    except Exception as e:
        print(f'{m} 失败: {str(e)[:150]}')
        continue

time.sleep(4)
try:
    c = doc.Tree.FindNode(r'\Data\Components\Specifications\Input')
    print('组分节点元素数:', c.Elements.Count)
    print('前 20 个子节点:', [c.Elements.Item(i).Name for i in range(min(20, c.Elements.Count))])
except Exception as e:
    print('读组分失败:', str(e)[:150])

# 关键验证：能否写流股温度
try:
    p = r'\Data\Streams\S-101\Input\TEMP'
    doc.Tree.FindNode(p).Value = 25.0
    print('>>> 流股温度写入成功:', doc.Tree.FindNode(p).Value)
except Exception as e:
    print('>>> 流股温度写入失败:', str(e)[:150])
