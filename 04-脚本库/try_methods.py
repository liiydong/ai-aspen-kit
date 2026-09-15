# -*- coding: utf-8 -*-
"""换物性方法（不需要二元参数的状态方程法）再运行"""
import time
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\noair2.bkp'
METHODS = ['PENG-ROB', 'RK-SOAVE', 'PR-BM']

for meth in METHODS:
    print('=' * 20, meth, flush=True)
    doc = win32.DispatchEx('Apwn.Document')
    try:
        doc.SuppressDialogs = True
    except Exception:
        pass
    doc.InitFromArchive(SRC)
    time.sleep(4)
    try:
        n = doc.Tree.FindNode(r'\Data\Properties\Specifications\Input\GOPSETNAME')
        n.Value = meth
        print('  物性方法设为 %r -> 读回 %r' % (meth, doc.Tree.FindNode(
            r'\Data\Properties\Specifications\Input\GOPSETNAME').Value), flush=True)
    except Exception as e:
        print('  设物性方法失败', str(e)[:100], flush=True)
    t0 = time.time()
    try:
        doc.Engine.Run2()
    except Exception as e:
        print('  Run2 异常', str(e)[:100], flush=True)
    while time.time() - t0 < 120:
        time.sleep(3)
        try:
            if doc.Engine.IsRunning is False:
                break
        except Exception:
            break
    print('  运行用时 %.0fs' % (time.time() - t0), flush=True)
    got = []
    for b in ['M-101', 'E-101', 'R-101', 'T-201', 'T-404', 'R-501', 'E-601']:
        n = doc.Tree.FindNode(r'\Data\Blocks\%s\Output\TEMP' % b)
        got.append('%s=%s' % (b, n.Value if n is not None else 'None'))
    print('  ', ' '.join(got), flush=True)
    # 流股输出
    s = doc.Tree.FindNode(r'\Data\Streams\S-104\Output\TEMP')
    print('   S-104.Out.TEMP =', s.Value if s is not None else None, flush=True)
    try:
        doc.SaveAs(r'D:\<化工工作区>\_probe\m_%s.bkp' % meth.replace('-', '_'))
    except Exception:
        pass
    try:
        doc.Close()
    except Exception:
        pass
    print(flush=True)
