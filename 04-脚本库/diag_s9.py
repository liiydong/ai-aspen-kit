# -*- coding: utf-8 -*-
"""诊断 s9 的 SEVERE ERROR 原文"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

MSGS = []


class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            MSGS.append(s)


doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
win32.WithEvents(doc, Sink)
doc.InitFromArchive2(r'D:\<化工工作区>\NA-Chemical-10000t_s9.bkp')
time.sleep(3)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 540:
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

L = ['=== 全部面板消息 (%d 条) ===' % len(MSGS)]
# 打印第一个 SEVERE ERROR 前后各 12 条
idx = [i for i, x in enumerate(MSGS) if 'SEVERE ERROR' in x]
L.append('SEVERE ERROR 出现在: %s' % idx[:5])
if idx:
    a = max(0, idx[0] - 12)
    b = min(len(MSGS), idx[0] + 14)
    for i in range(a, b):
        L.append('%3d | %s' % (i, MSGS[i][:170]))
L.append('')
L.append('=== 末尾 25 条 ===')
for i in range(max(0, len(MSGS) - 25), len(MSGS)):
    L.append('%3d | %s' % (i, MSGS[i][:170]))

open(r'D:\<化工工作区>\_probe\s9err.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('DONE', len(MSGS))
try:
    doc.Close()
except Exception:
    pass
