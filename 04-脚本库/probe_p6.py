# -*- coding: utf-8 -*-
"""核对 Aspen 实际读到的 T-401 参数（D:F / PRES1 / 塔顶温度/压力）"""
import sys, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32

for tag, P in [('p6a(D:F=0.659)', r'D:\<化工工作区>\_probe\p6a.bkp'),
               ('v19(D:F=0.995)', r'D:\<化工工作区>\NA-Chemical-10000t_v19.bkp')]:
    doc = win32.DispatchEx('Apwn.Document')
    doc.SuppressDialogs = True
    doc.InitFromArchive2(P)
    time.sleep(3)

    def g(p):
        n = doc.Tree.FindNode(p)
        if n is None:
            return 'NONE'
        try:
            return n.Value
        except Exception as ex:
            return 'ERR'

    print('=' * 60)
    print(tag)
    for f in ['D:F', 'BASIS_RR', 'PRES1', 'PRES_STAGE1', 'NSTAGE', 'BASIS_RDV']:
        print('  T-401 Input\\%-12s = %s' % (f, g(r'\Data\Blocks\T-401\Input\%s' % f)))
    bb = r'\Data\Blocks\T-401\Output'
    print('  输出: Ttop=%s Tbot=%s RR=%s Ptop=%s' % (
        g(bb + r'\TOP_TEMP'), g(bb + r'\BOTTOM_TEMP'), g(bb + r'\MOLE_RR'),
        g(bb + r'\TOP_PRES')))
    # 塔顶流股
    print('  S-115 TOL=%s  3-CP=%s' % (
        g(r'\Data\Streams\S-115\Output\MASSFLMX\MIXED'),
        g(r'\Data\Streams\S-115\Output\MASSFLOW3')))
    try:
        doc.Close()
    except Exception:
        pass
print('DONE')
