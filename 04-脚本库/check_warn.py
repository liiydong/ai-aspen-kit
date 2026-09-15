# -*- coding: utf-8 -*-
"""重跑用户修改后的模型，导出全部控制面板消息（含警告原文），并读取 E-104/E-201 等参数"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_BIP2.bkp'
MSGS = []


class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            MSGS.append(s)


doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
win32.WithEvents(doc, Sink)
doc.InitFromArchive2(SRC)
time.sleep(3)
print('\\Data children =', doc.Tree.FindNode(r'\Data').Elements.Count)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 600:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 10:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(80):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)


def g(p):
    n = doc.Tree.FindNode(p)
    try:
        return n.Value if n is not None else None
    except Exception:
        return None


print()
print('=== 控制面板消息总数: %d ===' % len(MSGS))
print()
print('--- 错误/警告汇总行 ---')
for i, m in enumerate(MSGS):
    if re.search(r'Terminal Errors|Severe Errors|Errors\s|Warnings|Summary of', m):
        print('%4d | %s' % (i, m[:150]))

print()
print('--- 全部 WARNING / ERROR 段落 ---')
for i, m in enumerate(MSGS):
    if re.search(r'\*\s*(WARNING|ERROR)|ERROR IN|CANNOT|IS NOT|MISSING|ZERO|NOT SPECIFIED', m, re.I):
        print('%4d | %s' % (i, m[:170]))

print()
print('--- 设备参数抽查 ---')
for b, keys in [('E-104', ['PRES', 'TEMP', 'DELP', 'SPEC_OPT']),
                ('E-201', ['PRES', 'TEMP', 'DELP', 'SPEC_OPT']),
                ('E-101', ['PRES', 'TEMP']), ('E-103', ['PRES', 'TEMP']),
                ('E-105', ['PRES', 'TEMP']), ('E-601', ['PRES', 'TEMP'])]:
    row = []
    for k in keys:
        for path in [r'\Data\Blocks\%s\Input\%s' % (b, k),
                     r'\Data\Blocks\%s\Input\PARAM\%s' % (b, k)]:
            v = g(path)
            if v is not None:
                row.append('%s=%s' % (k, v))
                break
    print('  %-7s %s' % (b, ' | '.join(row)))

print()
print('--- 物料平衡 ---')
IN = ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-132', 'S-210']
OUTL = ['S-130', 'S-131', 'S-211', 'S-212', 'S-213', 'S-115', 'S-119',
        'S-122', 'S-216', 'S-217', 'S-218', 'S-405C', 'S-406', 'S-209']
vi = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in IN)
vo = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in OUTL)
print('  进=%.2f 出=%.2f 偏差=%.4f%%' % (vi, vo, abs(vi - vo) / vi * 100 if vi else 0))
print('  产品 S-209 =', g(r'\Data\Streams\S-209\Output\MASSFLMX\MIXED'))

try:
    doc.Close()
except Exception:
    pass
try:
    doc.Quit()
except Exception:
    pass
print('DONE')
