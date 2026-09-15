# -*- coding: utf-8 -*-
"""修 RUN-CLASS + 清理 CID 名，验证 18 组分与流程图是否都正常"""
import time, re
import win32com.client as win32

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_修复.bkp'
A = r'D:\<化工工作区>\_probe\wa.bkp'   # 只修 RUN-CLASS
B = r'D:\<化工工作区>\_probe\wb.bkp'   # 再清理 CID

t0 = open(SRC, encoding='utf-8', errors='ignore').read()

tA = t0.replace('RUN-CLASS = PROP', 'RUN-CLASS = FLOWSHEET')
open(A, 'w', encoding='utf-8', errors='ignore').write(tA)

REN = [('4-MET-01', 'NA2SO4'), ('4-MET-02', '4-MP'), ('NICOT-01', 'NAM'),
       ('NIACI-01', 'NAC'), ('SODIU-01', 'NAOH'), ('SULFU-01', 'H2SO4')]
tB = tA
for old, new in REN:
    n = tB.count('"%s"' % old)
    tB = tB.replace('"%s"' % old, '"%s"' % new)
    print('  %-10s -> %-8s 替换 %d 处' % (old, new, n))
open(B, 'w', encoding='utf-8', errors='ignore').write(tB)
print()

BASE = ['3-MP', 'NH3', 'O2', 'N2', '3-CP', '4-CP', 'H2O', 'TOL', 'CO2', 'HCN', 'CO', 'AIR']


def check(path, tag):
    print('=' * 18, tag, flush=True)
    doc = win32.DispatchEx('Apwn.Document')
    try:
        doc.SuppressDialogs = True
    except Exception:
        pass
    doc.InitFromArchive(path)
    time.sleep(4)
    d = doc.Tree.FindNode(r'\Data')
    bl = doc.Tree.FindNode(r'\Data\Blocks')
    st = doc.Tree.FindNode(r'\Data\Streams')
    ms = doc.Tree.FindNode(r'\Data\Properties\Molecular Structure')
    names = [ms.Elements.Item(i).Name for i in range(ms.Elements.Count)] if ms else []
    print('  \\Data=%s 模块=%s 流股=%s 组分=%d' % (
        d.Elements.Count if d else 0,
        bl.Elements.Count if bl else 0,
        st.Elements.Count if st else 0, len(names)), flush=True)
    print('  组分:', names, flush=True)
    print('  新增:', [x for x in names if x not in BASE], flush=True)
    try:
        doc.Close()
    except Exception:
        pass
    print()


check(A, 'A: 只修 RUN-CLASS')
check(B, 'B: RUN-CLASS + CID 改名')
