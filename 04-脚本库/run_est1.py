# -*- coding: utf-8 -*-
"""估算测试（默认 UNIFAC）"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>'
P = os.path.join(BASE, 'NA_est_dmd.bkp')

msgs = []
class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            msgs.append(s)

doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
win32.WithEvents(doc, Sink)
doc.InitFromArchive2(P)
t = time.time()
while time.time() - t < 25:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)

n0 = len(msgs)
doc.Engine.Run2(False)
t = time.time()
while time.time() - t < 360:
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
out = []
out.append('新消息 %d 条' % len(new))
fg = [m for m in new if 'FUNCTIONAL GROUP' in m or 'NOT MATCHED' in m or 'CANNOT BE ESTIMATED' in m]
out.append('基团/估算相关 %d 条:' % len(fg))
for m in fg[:8]:
    out.append('  | ' + m[:120])
for m in new[-6:]:
    out.append('  > ' + m[:110])

out.append('--- 参数状态 ---')
for k in ['Input\\VAL1', 'Output\\CID1', 'Output\\VALUE']:
    n = doc.Tree.FindNode(r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\%s' % k)
    try:
        out.append('  NRTL-1\\%-16s rows=%s' % (k, n.Elements.Count))
    except Exception:
        out.append('  NRTL-1\\%-16s ?' % k)
n = doc.Tree.FindNode(r'\Data\Components\UNIFAC-Groups\Input\GROUPNO')
try:
    out.append('  UNIFAC-Groups\\GROUPNO rows=%s' % n.Elements.Count)
except Exception:
    out.append('  UNIFAC-Groups\\GROUPNO ?')

print('\n'.join(out))
try:
    doc.Close()
except Exception:
    pass
print('DONE')
