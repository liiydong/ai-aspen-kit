# -*- coding: utf-8 -*-
"""读各模块 Output 里的 ERR / PER_ERROR，找报错"""
import time
import win32com.client as win32

F = r'D:\<化工工作区>\_probe\full3.bkp'
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(F)
time.sleep(4)
print('打开 Ready =', doc.Engine.Ready, flush=True)
try:
    doc.Engine.Run2()
except Exception as e:
    print('Run2 异常', str(e)[:120], flush=True)
t0 = time.time()
while time.time() - t0 < 90:
    time.sleep(3)
    try:
        if doc.Engine.IsRunning is False:
            break
    except Exception:
        break
print('运行用时 %.0fs' % (time.time() - t0), flush=True)
print()
BLOCKS = ['M-101', 'E-101', 'R-101', 'E-104', 'T-201', 'E-201', 'T-301',
          'T-401', 'T-402', 'T-403', 'T-404', 'R-501', 'F-501', 'E-601']
for b in BLOCKS:
    n = doc.Tree.FindNode(r'\Data\Blocks\%s\Output' % b)
    if n is None:
        print('  %-7s 无 Output' % b, flush=True)
        continue
    info = []
    for f in ['ERR', 'PER_ERROR', 'TEMP', 'PRES', 'DUTY', 'MASS_RR', 'MOLE_RR']:
        try:
            x = n.Elements.Item(f)
            if x is None:
                continue
            v = x.Value
            if v is not None and v != '' and v != 0:
                info.append('%s=%r' % (f, v))
        except Exception:
            pass
    print('  %-7s %s' % (b, '; '.join(info) if info else '(全为空)'), flush=True)

print()
print('=== 全局搜索报错文本节点 ===', flush=True)
for p in [r'\Data\Results Summary\Run-Status\Error',
          r'\Data\Results Summary\Run-Status\Message',
          r'\Data\Results Summary\Stream-Sum\ERR']:
    n = doc.Tree.FindNode(p)
    print('  %-46s %r' % (p, n.Value if n is not None else None), flush=True)
try:
    doc.SaveAs(r'D:\<化工工作区>\_probe\full3_run.bkp')
except Exception:
    pass
try:
    doc.Close()
except Exception:
    pass
