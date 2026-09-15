# -*- coding: utf-8 -*-
"""dump r3.bkp 的完整控制面板消息（找出具体块错误）"""
import sys, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

P = r'D:\<化工工作区>\NA-Chemical-10000t_r3.bkp'
MSGS = []


class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            MSGS.append(s)


doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
win32.WithEvents(doc, Sink)
doc.InitFromArchive2(P)
time.sleep(3)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 300:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 10:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(50):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)
print('消息数:', len(MSGS))
for i, x in enumerate(MSGS[:70], 1):
    print('%3d| %s' % (i, x[:170]))
try:
    doc.Close()
except Exception:
    pass
print('DONE')
