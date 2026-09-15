# -*- coding: utf-8 -*-
"""测试：1) DBNAME1-only 能否解析 4 个组分；2) InitFromArchive3 能否加载流程图"""
import time, sys, shutil, re
import win32com.client as win32

BASE = r'D:\<化工工作区>\_probe\base.bkp'
P = r'D:\<化工工作区>\_probe\testP.bkp'
Q = r'D:\<化工工作区>\_probe\testQ.bkp'

text = open(BASE, encoding='utf-8', errors='ignore').read()
start = text.find('? COMPONENTS MAIN ?')
marker = '\\ ? COMPONENTS "ADA/PCS"'
end = text.find(marker, start)


def make(path, entries):
    add = ' /  ' + ' /  '.join(entries) + ' '
    open(path, 'w', encoding='utf-8', errors='ignore').write(text[:end] + add + text[end:])


# P: 只给 DBNAME1，不给别名
make(P, [
    'CID = "NAC" OUTNAME = "NAC" TYPE = CONV DBNAME1 = "NICOTINIC-ACID"',
    'CID = "NAOHC" OUTNAME = "NAOHC" TYPE = CONV DBNAME1 = "SODIUM-HYDROXIDE"',
    'CID = "H2SO4C" OUTNAME = "H2SO4C" TYPE = CONV DBNAME1 = "SULFURIC-ACID"',
    'CID = "NA2SO4C" OUTNAME = "NA2SO4C" TYPE = CONV DBNAME1 = "SODIUM-SULFATE"',
])

# Q: 常用式作别名
make(Q, [
    'CID = "NAC2" ANAME = C6H5NO2-D1 OUTNAME = "NAC2" TYPE = CONV DBNAME1 = "NICOTINIC-ACID" ANAME1 = "C6H5NO2-D1"',
    'CID = "NAOH2" ANAME = NAOH OUTNAME = "NAOH2" TYPE = CONV DBNAME1 = "SODIUM-HYDROXIDE" ANAME1 = "NAOH"',
    'CID = "H2SO42" ANAME = H2SO4 OUTNAME = "H2SO42" TYPE = CONV DBNAME1 = "SULFURIC-ACID" ANAME1 = "H2SO4"',
    'CID = "NA2SO42" ANAME = NA2SO4 OUTNAME = "NA2SO42" TYPE = CONV DBNAME1 = "SODIUM-SULFATE" ANAME1 = "NA2SO4"',
])
print('测试文件已生成', flush=True)


def probe(path, tag, method='InitFromArchive'):
    print('=' * 20, tag, method, flush=True)
    doc = win32.DispatchEx('Apwn.Document')
    try:
        doc.SuppressDialogs = True
    except Exception:
        pass
    try:
        getattr(doc, method)(path)
    except Exception as e:
        print('  打开失败:', str(e)[:150], flush=True)
        return
    time.sleep(3)
    ms = doc.Tree.FindNode(r'\Data\Properties\Molecular Structure')
    if ms is not None:
        names = [ms.Elements.Item(i).Name for i in range(ms.Elements.Count)]
        print('  组分 %d: %s' % (len(names), names), flush=True)
    else:
        print('  组分: None', flush=True)
    for p in [r'\Data\Streams', r'\Data\Blocks']:
        n = doc.Tree.FindNode(p)
        print('  %-16s %s' % (p, 'None' if n is None else '存在(%s)' % n.Elements.Count), flush=True)
    try:
        doc.Close()
    except Exception:
        pass
    print(flush=True)


probe(P, 'P: DBNAME1-only')
probe(Q, 'Q: 常用式别名')
probe(BASE, 'base: InitFromArchive3', 'InitFromArchive3')
