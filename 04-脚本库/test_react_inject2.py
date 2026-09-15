# -*- coding: utf-8 -*-
"""重试：修正反斜杠与单位标记"""
import time
import win32com.client as win32

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_工作版.bkp'
OUT = r'D:\<化工工作区>\_probe\react_test2.bkp'

t = open(SRC, encoding='utf-8', errors='ignore').read()

SEC = ('? BLOCK RSTOIC "R-501" ? ; "METCBAR_MOLE" ; ; ICON1 ; '
       '\\ PARAM TEMP = 140.0 PRES = 4.0 SPEC-OPT = TP '
       '\\ \\ STOIC REACNO = 1 STOIC-CID = "3-CP" STOIC-SSID = MIXED COEF = -1.0 <0> <0> '
       '\\ \\ STOIC REACNO = 1 STOIC-CID = H2O STOIC-SSID = MIXED COEF = -1.0 <0> <0> '
       '\\ \\ STOIC1 REACNO1 = 1 STOIC-CID1 = NAM STOIC-SSID1 = MIXED COEF1 = 1.0 <0> <0> '
       '\\ \\ CONVEX EXT-REACNO = 1 KEY-SSID = MIXED KEY-CID = "3-CP" CONV = .99 <0> <0> '
       '\\ \\ PRODUCTS SID = "S-201" \\ ')

anchor = '? BLOCK RSTOIC "F-501" ?'
i = t.find(anchor)
t2 = t[:i] + SEC + '\n' + t[i:]
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t2)
print('已写', OUT, flush=True)
print('注入片段:', repr(SEC[SEC.find('PARAM'):][:160]), flush=True)

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(OUT)
time.sleep(4)

rx = doc.Tree.FindNode(r'\Data\Reactions')
print('Reactions 子节点:', [rx.Elements.Item(i2).Name for i2 in range(rx.Elements.Count)], flush=True)
rr = doc.Tree.FindNode(r'\Data\Reactions\Reactions')
print('Reactions\\Reactions 数:', rr.Elements.Count if rr is not None else 'None', flush=True)
if rr is not None and rr.Elements.Count:
    for i2 in range(rr.Elements.Count):
        print('   反应集:', rr.Elements.Item(i2).Name, flush=True)
n = doc.Tree.FindNode(r'\Data\Blocks\R-501\Input')
if n is not None:
    for f in ['TEMP', 'PRES']:
        nn = doc.Tree.FindNode(r'\Data\Blocks\R-501\Input\\' + f)
        print('   %s = %r' % (f, nn.Value if nn is not None else None), flush=True)
    # 找反应集引用
    for f in ['REACSET', 'REACTIONS', 'RXNID', 'REACID']:
        nn = doc.Tree.FindNode(r'\Data\Blocks\R-501\Input\\' + f)
        print('   %-10s %s' % (f, repr(nn.Value) if nn is not None else 'None'), flush=True)
try:
    doc.Close()
except Exception:
    pass
