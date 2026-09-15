# -*- coding: utf-8 -*-
"""深挖 COEF 子树结构 + 试后缀写法"""
import time
import win32com.client as win32

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_工作版.bkp'

VARIANTS = {
    'X_后缀': ('? BLOCK RSTOIC "R-501" ? ; "METCBAR_MOLE" ; ; ICON1 ; '
               '\\ PARAM TEMP = 140.0 PRES = 4.0 SPEC-OPT = TP '
               '\\ STOIC REACNO = 1 STOIC-CID = "3-CP" STOIC-SSID = MIXED COEF = -1.0 <0> <0> '
               '\\ STOIC1 REACNO1 = 1 STOIC-CID1 = H2O STOIC-SSID1 = MIXED COEF1 = -1.0 <0> <0> '
               '\\ STOIC2 REACNO2 = 1 STOIC-CID2 = NAM STOIC-SSID2 = MIXED COEF2 = 1.0 <0> <0> '
               '\\ CONVEX EXT-REACNO = 1 KEY-SSID = MIXED KEY-CID = "3-CP" CONV = .99 <0> <0> '
               '\\ PRODUCTS SID = "S-201" \\ '),
}

anchor = '? BLOCK RSTOIC "F-501" ?'


def dump(node, path, depth=0, maxd=3):
    if node is None or depth > maxd:
        return
    try:
        c = node.Elements.Count
    except Exception:
        try:
            print('   ' + '  ' * depth + '%s = %r' % (path, node.Value))
        except Exception:
            pass
        return
    try:
        v = repr(node.Value)[:40]
    except Exception:
        v = ''
    print('   ' + '  ' * depth + '%s  [%d] %s' % (path, c, v))
    for i in range(min(c, 8)):
        ch = node.Elements.Item(i)
        dump(ch, ch.Name, depth + 1, maxd)


for tag, sec in VARIANTS.items():
    out = r'D:\<化工工作区>\_probe\var_%s.bkp' % tag
    t = open(SRC, encoding='utf-8', errors='ignore').read()
    i = t.find(anchor)
    open(out, 'w', encoding='utf-8', errors='ignore').write(t[:i] + sec + '\n' + t[i:])
    print('=' * 20, tag, flush=True)
    doc = win32.DispatchEx('Apwn.Document')
    try:
        doc.SuppressDialogs = True
    except Exception:
        pass
    doc.InitFromArchive(out)
    time.sleep(4)
    base = r'\Data\Blocks\R-501\Input'
    for f in ['TEMP', 'PRES', 'COEF', 'CONV', 'KEY_CID', 'EXTENT']:
        n = doc.Tree.FindNode(base + '\\' + f)
        print('  == %s' % f, flush=True)
        dump(n, f, 0, 4)
    try:
        doc.Close()
    except Exception:
        pass
    print(flush=True)
