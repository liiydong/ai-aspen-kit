# -*- coding: utf-8 -*-
"""测试：向 bkp 注入 RStoic 反应段，Aspen 是否接受"""
import time, re
import win32com.client as win32

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_工作版.bkp'
OUT = r'D:\<化工工作区>\_probe\react_test.bkp'

t = open(SRC, encoding='utf-8', errors='ignore').read()

# R-501 反应段（照 pfdtut.bkp 的 RSTOIC 格式）
SEC = ('? BLOCK RSTOIC "R-501" ? ; "METCBAR_MOLE" ; ; ICON1 ; \\ '
       'PARAM TEMP = 140.0 PRES = 4.0 SPEC-OPT = TP \\ '
       '\\ STOIC REACNO = 1 STOIC-CID = "3-CP" STOIC-SSID = MIXED COEF = -1.0 \\ '
       '\\ STOIC REACNO = 1 STOIC-CID = H2O STOIC-SSID = MIXED COEF = -1.0 \\ '
       '\\ STOIC1 REACNO1 = 1 STOIC-CID1 = NAM STOIC-SSID1 = MIXED COEF1 = 1.0 \\ '
       '\\ CONVEX EXT-REACNO = 1 KEY-SSID = MIXED KEY-CID = "3-CP" CONV = .99 \\ '
       '\\ PRODUCTS SID = "S-201" \\ ')

anchor = '? BLOCK RSTOIC "F-501" ?'
i = t.find(anchor)
print('锚点位置:', i)
if i < 0:
    raise SystemExit('未找到锚点')
t2 = t[:i] + SEC + '\n' + t[i:]
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t2)
print('已写出', OUT, flush=True)

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(OUT)
time.sleep(4)

d = doc.Tree.FindNode(r'\Data')
print('\\Data 子节点:', d.Elements.Count if d else 0, flush=True)
bl = doc.Tree.FindNode(r'\Data\Blocks')
print('模块数:', bl.Elements.Count if bl else 0, flush=True)

rx = doc.Tree.FindNode(r'\Data\Reactions\Reactions')
print('Reactions/Reactions 子节点:', rx.Elements.Count if rx is not None else 'None', flush=True)
if rx is not None and rx.Elements.Count:
    for i2 in range(rx.Elements.Count):
        print('   ', rx.Elements.Item(i2).Name, flush=True)

# R-501 的输入字段里有没有反应相关
n = doc.Tree.FindNode(r'\Data\Blocks\R-501\Input')
if n is not None:
    names = [n.Elements.Item(i2).Name for i2 in range(n.Elements.Count)]
    rel = [x for x in names if any(k in x.upper() for k in
           ['STOIC', 'COEF', 'CONV', 'REAC', 'EXTENT', 'KEY'])]
    print('R-501 反应相关字段:', rel, flush=True)
    for f in ['TEMP', 'PRES', 'SPEC_OPT']:
        p2 = r'\Data\Blocks\R-501\Input' + '\\' + f
        nn = doc.Tree.FindNode(p2)
        print('   %-10s = %r' % (f, nn.Value if nn is not None else 'None'), flush=True)
try:
    doc.Close()
except Exception:
    pass
