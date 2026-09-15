# -*- coding: utf-8 -*-
"""最后尝试：触发 Aspen 从数据库加载分子结构"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>'
MS = r'\Data\Properties\Molecular Structure'

msgs = []


class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            msgs.append(s)


doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
win32.WithEvents(doc, Sink)
doc.InitFromArchive2(os.path.join(BASE, 'NA_work.bkp'))
t = time.time()
while time.time() - t < 25:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)

print('--- 尝试 A：往 H2O 分子结构列 Add 行 ---')
for col in ['ATOMTYPE', 'ATOMNUM', 'BONDTYPE']:
    nd = doc.Tree.FindNode(MS + r'\H2O\Input\%s' % col)
    if nd is None:
        print('  %s -> None' % col); continue
    try:
        print('  %s rows=%d' % (col, nd.Elements.Count))
    except Exception as ex:
        print('  %s ?' % col)
# 试 Add
for path, val in [(MS + r'\H2O\Input\ATOMTYPE', 'O'),
                  (MS + r'\H2O\Input\ATOMNUM', 1)]:
    nd = doc.Tree.FindNode(path)
    try:
        nd.Elements.Add(val)
        print('  Add %s=%r 成功' % (path.split('\\')[-1], val))
    except Exception as ex:
        print('  Add %s=%r 失败: %s' % (path.split('\\')[-1], val, str(ex)[:80]))

print()
print('--- 尝试 B：设 CC Nodes 标志 ---')
for k, v in [('ASKED', 1), ('HASFORMUL', 1), ('GEN', 1), ('FORMULA', 1)]:
    nd = doc.Tree.FindNode(MS + r'\H2O\CC Nodes\%s' % k)
    if nd is None:
        print('  %s None' % k); continue
    try:
        old = nd.Value
        nd.Value = v
        print('  %s: %r -> %r' % (k, old, nd.Value))
    except Exception as ex:
        print('  %s 失败: %s' % (k, str(ex)[:70]))

print()
print('--- 尝试 C：Reinit + Run ---')
try:
    doc.Engine.Reinit()
    print('  Reinit ok')
except Exception as ex:
    print('  Reinit:', str(ex)[:70])

n0 = len(msgs)
doc.Engine.Run2(False)
t = time.time()
while time.time() - t < 300:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.3)
    if time.time() - t > 10:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(100):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)
new = msgs[n0:]
print('  新消息:', len(new))
unif = [m for m in new if 'UNIFAC' in m or 'MISSING' in m]
print('  UNIFAC/MISSING 相关:', len(unif))

print()
print('--- 检查结果 ---')
for p in [MS + r'\H2O\Input\ATOMTYPE',
          MS + r'\H2O\Input\ATOMNUM',
          MS + r'\H2O\Input\BONDTYPE',
          r'\Data\Components\UNIFAC-Groups\Input\GROUPNO',
          r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Output\CID1']:
    nd = doc.Tree.FindNode(p)
    if nd is None:
        print('  %-58s None' % p); continue
    try:
        print('  %-58s rows=%d' % (p, nd.Elements.Count))
    except Exception:
        try:
            print('  %-58s val=%r' % (p, nd.Value))
        except Exception:
            print('  %-58s ?' % p)

try:
    doc.Close()
except Exception:
    pass
print('DONE')
