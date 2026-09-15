# -*- coding: utf-8 -*-
"""读 Aspen 控制面板的真实报错"""
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

cp = doc.Engine.ControlPanel
print('ControlPanel 类型:', type(cp))
print('可用成员:', [m for m in dir(cp) if not m.startswith('_')][:30], flush=True)
print()
try:
    print('Value:', repr(cp.Value)[:2000])
except Exception as e:
    print('Value 读取失败:', str(e)[:120])

print()
print('=== 运行 ===')
try:
    doc.Engine.Run2()
    print('Run2 完成', flush=True)
except Exception as e:
    print('Run2 抛出:', str(e)[:150])
time.sleep(2)
print()
print('=== 运行后的控制面板 ===')
try:
    cp = doc.Engine.ControlPanel
    v = cp.Value
    print(str(v)[:4000])
except Exception as e:
    print('读取失败:', str(e)[:150])
try:
    doc.Close()
except Exception:
    pass
