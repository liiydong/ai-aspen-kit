# -*- coding: utf-8 -*-
"""用单反斜杠格式注入，检查 COEF/CONV/KEY_CID 表格是否被填充"""
import time
import win32com.client as win32

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_工作版.bkp'
OUT = r'D:\<化工工作区>\_probe\react_test3.bkp'

t = open(SRC, encoding='utf-8', errors='ignore').read()

SEC = ('? BLOCK RSTOIC "R-501" ? ; "METCBAR_MOLE" ; ; ICON1 ; '
       '\\ PARAM TEMP = 140.0 PRES = 4.0 SPEC-OPT = TP '
       '\\ STOIC REACNO = 1 STOIC-CID = "3-CP" STOIC-SSID = MIXED COEF = -1.0 <0> <0> '
       '\\ STOIC REACNO = 1 STOIC-CID = H2O STOIC-SSID = MIXED COEF = -1.0 <0> <0> '
       '\\ STOIC REACNO = 1 STOIC-CID = NAM STOIC-SSID = MIXED COEF = 1.0 <0> <0> '
       '\\ CONVEX EXT-REACNO = 1 KEY-SSID = MIXED KEY-CID = "3-CP" CONV = .99 <0> <0> '
       '\\ PRODUCTS SID = "S-201" \\ ')

anchor = '? BLOCK RSTOIC "F-501" ?'
i = t.find(anchor)
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t[:i] + SEC + '\n' + t[i:])
print('注入:', repr(SEC[SEC.find('PARAM'):][:120]), flush=True)

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(OUT)
time.sleep(4)

base = r'\Data\Blocks\R-501\Input'
for f in ['TEMP', 'PRES', 'SPEC_OPT', 'COEF', 'COEF1', 'CONV', 'CONVERSION',
          'EXTENT', 'KEY_CID', 'KEY_SSID', 'PROD_NOX']:
    n = doc.Tree.FindNode(base + '\\' + f)
    if n is None:
        print('  %-12s None' % f, flush=True)
        continue
    try:
        c = n.Elements.Count
        kids = [n.Elements.Item(i2).Name for i2 in range(min(6, c))]
        print('  %-12s 子=%s %s' % (f, c, kids), flush=True)
    except Exception:
        print('  %-12s 值=%r' % (f, n.Value), flush=True)
try:
    doc.Close()
except Exception:
    pass
