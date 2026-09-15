# -*- coding: utf-8 -*-
"""确认 Aspen 运行是否异步：Run2 后轮询状态，等真正跑完再看结果"""
import time
import win32com.client as win32

F = r'D:\<化工工作区>\NA-Chemical-10000t_模拟.bkp'
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(F)
time.sleep(4)

E = doc.Engine
print('运行前: Ready=%r  IsRunning=%r' % (E.Ready, E.IsRunning), flush=True)
try:
    print('RunControl =', E.RunControl, flush=True)
except Exception as e:
    print('RunControl 读取失败', str(e)[:60], flush=True)
try:
    print('RunSettings =', E.RunSettings, flush=True)
except Exception as e:
    print('RunSettings 读取失败', str(e)[:60], flush=True)

print()
print('=== 启动运行 ===', flush=True)
t0 = time.time()
try:
    E.Run2()
    print('Run2 已调用 (%.1fs)' % (time.time() - t0), flush=True)
except Exception as e:
    print('Run2 异常', str(e)[:120], flush=True)

# 轮询
for i in range(40):
    time.sleep(3)
    try:
        run = E.IsRunning
    except Exception as e:
        run = 'ERR:%s' % str(e)[:40]
    try:
        ready = E.Ready
    except Exception:
        ready = '?'
    if i % 3 == 0 or run is False:
        print('  [%5.1fs] IsRunning=%r Ready=%r' % (time.time() - t0, run, ready), flush=True)
    if run is False and i > 1:
        break

print()
print('=== 结果检查 ===', flush=True)
for s in ['S-104', 'S-106', 'S-110', 'S-201', 'S-301']:
    n = doc.Tree.FindNode(r'\Data\Streams\%s\Output\TEMP' % s)
    v = n.Value if n is not None else 'NONE'
    print('  %-8s Out.TEMP = %r' % (s, v), flush=True)
for b in ['R-101', 'R-501']:
    for f in ['TEMP', 'PRES']:
        n = doc.Tree.FindNode(r'\Data\Blocks\%s\Output\%s' % (b, f))
        print('  %s.Out.%s = %r' % (b, f, n.Value if n is not None else 'NONE'), flush=True)
try:
    doc.SaveAs(r'D:\<化工工作区>\_probe\ran.bkp')
    print('已保存 ran.bkp', flush=True)
except Exception as e:
    print('SaveAs 失败', str(e)[:80], flush=True)
try:
    doc.Close()
except Exception:
    pass
