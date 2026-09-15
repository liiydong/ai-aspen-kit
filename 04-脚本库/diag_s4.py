# -*- coding: utf-8 -*-
"""定位 s3 的 input specification 报错：导出全部控制面板消息"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_s3.bkp'

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
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 480:
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

lines = []
lines.append('=== 控制面板全部消息 (%d 条) ===' % len(MSGS))
for i, x in enumerate(MSGS):
    lines.append('%3d | %s' % (i, x))


def g(p):
    n = doc.Tree.FindNode(p)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return None


lines.append('')
lines.append('=== 各模块是否在流程中 ===')
for b in ['M-101', 'E-101', 'E-103', 'E-105', 'R-101', 'E-104', 'B-101', 'B-102', 'B-103',
          'T-201', 'T-202', 'E-201', 'T-301', 'T-401', 'T-402', 'T-403', 'T-404',
          'R-501', 'T-302', 'E-501', 'V-501', 'E-601', 'E-602', 'E-603', 'C-601', 'M-601', 'D-601']:
    v = g(r'\Data\Blocks\%s\Input\TYPE' % b)
    lines.append('  %-7s TYPE=%s' % (b, v))

try:
    doc.Close()
except Exception:
    pass

open(r'D:\<化工工作区>\_probe\s4diag.txt', 'w', encoding='utf-8').write('\n'.join(lines))
print('DONE', len(MSGS))
