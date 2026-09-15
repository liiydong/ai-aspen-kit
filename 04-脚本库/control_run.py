# -*- coding: utf-8 -*-
"""对照：参考模型能否用 COM 跑出结果"""
import time
import win32com.client as win32

REF = r'D:\<化工工作区>\_probe\ref_multi.bkp'
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(REF)
time.sleep(4)

print('打开后: Ready=%r IsRunning=%r' % (doc.Engine.Ready, doc.Engine.IsRunning), flush=True)
print('结果字段（打开时）:', flush=True)
for s in ['S5', 'S7']:
    n = doc.Tree.FindNode(r'\Data\Streams\%s\Output\TEMP' % s)
    print('   %s.Out.TEMP = %r' % (s, n.Value if n is not None else None), flush=True)

print()
print('=== 运行 ===', flush=True)
t0 = time.time()
try:
    doc.Engine.Run2()
    print('  Run2 调用完成 %.1fs' % (time.time() - t0), flush=True)
except Exception as e:
    print('  Run2 异常', str(e)[:120], flush=True)
for i in range(12):
    time.sleep(3)
    try:
        r = doc.Engine.IsRunning
    except Exception:
        r = '?'
    if i % 4 == 3:
        print('   [%4.0fs] IsRunning=%r Ready=%r' % (time.time() - t0, r, doc.Engine.Ready), flush=True)
    if r is False and i > 0:
        break

print()
print('=== 运行后结果 ===', flush=True)
for s in ['S5', 'S7', 'S4']:
    n = doc.Tree.FindNode(r'\Data\Streams\%s\Output\TEMP' % s)
    print('   %s.Out.TEMP = %r' % (s, n.Value if n is not None else None), flush=True)
for b in ['B1', 'B2']:
    n = doc.Tree.FindNode(r'\Data\Blocks\%s\Output\COND_DUTY' % b)
    print('   %s.Out.COND_DUTY = %r' % (b, n.Value if n is not None else None), flush=True)
try:
    doc.Close()
except Exception:
    pass
