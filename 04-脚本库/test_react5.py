# -*- coding: utf-8 -*-
"""验证 COEF 行数是否按反应数变化 + 设置物性方法"""
import time
import win32com.client as win32

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_工作版.bkp'
OUT = r'D:\<化工工作区>\_probe\react_two.bkp'

# R-501：主反应(3-CP+H2O→NAM) + 副反应(NAM+H2O→NAC+NH3)
SEC = ('? BLOCK RSTOIC "R-501" ? ; "METCBAR_MOLE" ; ; ICON1 ; '
       '\\ PARAM TEMP = 140.0 PRES = 4.0 SPEC-OPT = TP '
       '\\ STOIC REACNO = 1 STOIC-CID = "3-CP" STOIC-SSID = MIXED COEF = -1.0 <0> <0> '
       '\\ STOIC1 REACNO1 = 1 STOIC-CID1 = H2O STOIC-SSID1 = MIXED COEF1 = -1.0 <0> <0> '
       '\\ STOIC2 REACNO2 = 1 STOIC-CID2 = NAM STOIC-SSID2 = MIXED COEF2 = 1.0 <0> <0> '
       '\\ STOIC3 REACNO3 = 2 STOIC-CID3 = NAM STOIC-SSID3 = MIXED COEF3 = -1.0 <0> <0> '
       '\\ STOIC4 REACNO4 = 2 STOIC-CID4 = H2O STOIC-SSID4 = MIXED COEF4 = -1.0 <0> <0> '
       '\\ STOIC5 REACNO5 = 2 STOIC-CID5 = NAC STOIC-SSID5 = MIXED COEF5 = 1.0 <0> <0> '
       '\\ STOIC6 REACNO6 = 2 STOIC-CID6 = NH3 STOIC-SSID6 = MIXED COEF6 = 1.0 <0> <0> '
       '\\ CONVEX EXT-REACNO = 1 KEY-SSID = MIXED KEY-CID = "3-CP" CONV = .985 <0> <0> '
       '\\ CONVEX1 EXT-REACNO1 = 2 KEY-SSID1 = MIXED KEY-CID1 = NAM CONV1 = .005 <0> <0> '
       '\\ PRODUCTS SID = "S-201" \\ ')

t = open(SRC, encoding='utf-8', errors='ignore').read()
i = t.find('? BLOCK RSTOIC "F-501" ?')
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t[:i] + SEC + '\n' + t[i:])
print('已写', OUT, flush=True)

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(OUT)
time.sleep(4)

base = r'\Data\Blocks\R-501\Input'
for f in ['TEMP', 'PRES', 'COEF', 'CONV', 'KEY_CID', 'EXTENT', 'REACNO', 'REACNO1']:
    n = doc.Tree.FindNode(base + '\\' + f)
    if n is None:
        print('  %-10s None' % f, flush=True)
        continue
    try:
        c = n.Elements.Count
        print('  %-10s 子=%s' % (f, c), flush=True)
    except Exception:
        print('  %-10s 值=%r' % (f, n.Value), flush=True)

print()
print('--- 设物性方法 NRTL ---', flush=True)
for path, val in [(r'\Data\Properties\Specifications\Input\GOPSETNAME', 'NRTL'),
                  (r'\Data\Properties\Specifications\Input\BASEOPSET', 'NRTL')]:
    n = doc.Tree.FindNode(path)
    if n is None:
        print('  %s None' % path, flush=True)
        continue
    try:
        n.Value = val
        print('  %-14s -> %r' % (path.split('\\')[-1],
              doc.Tree.FindNode(path).Value), flush=True)
    except Exception as e:
        print('  %-14s 失败 %s' % (path.split('\\')[-1], str(e)[:70]), flush=True)
try:
    doc.Close()
except Exception:
    pass
