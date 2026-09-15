# -*- coding: utf-8 -*-
"""功能验证：流股能否写入温度 + 新增组分（NAM/NAOH）的流量"""
import time, shutil, sys
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\x_NA-Chemical-10000t.bkp.backup'
BAS = r'D:\<化工工作区>\_probe\base.bkp'
MOD = r'D:\<化工工作区>\_probe\mod.bkp'


def setnode(doc, path, val):
    try:
        n = doc.Tree.FindNode(path)
        if n is None:
            return '节点不存在'
        n.Value = val
        return 'OK -> %s' % doc.Tree.FindNode(path).Value
    except Exception as e:
        return 'ERR %s' % str(e)[:90]


def probe(path, tag):
    print('=' * 18, tag)
    doc = win32.DispatchEx('Apwn.Document')
    try:
        doc.InitFromArchive2(path)
    except Exception as e:
        print('  打开失败:', str(e)[:120])
        return
    time.sleep(3)
    print('  TEMP 写入      :', setnode(doc, r'\Data\Streams\S-101\Input\TEMP', 25.0))
    print('  PRES 写入      :', setnode(doc, r'\Data\Streams\S-101\Input\PRES', 1.0))
    print('  NAM 摩尔流量   :', setnode(doc, r'\Data\Streams\S-101\Input\FLOW\MIXED\NAM', 1.0))
    print('  NAOH 摩尔流量  :', setnode(doc, r'\Data\Streams\S-101\Input\FLOW\MIXED\NAOH', 1.0))
    print('  3-MP 摩尔流量  :', setnode(doc, r'\Data\Streams\S-101\Input\FLOW\MIXED\3-MP', 1.0))
    # 保存成新 bkp 供查证
    out = path.replace('.bkp', '_saved.bkp')
    try:
        doc.SaveAs(out)
        time.sleep(2)
        print('  已另存:', out)
    except Exception as e:
        print('  SaveAs 失败:', str(e)[:100])
    try:
        doc.Close()
    except Exception:
        pass
    print()


probe(BAS, '[基线] 12 组分')
probe(MOD, '[修改] 18 组分')
