# -*- coding: utf-8 -*-
"""保存提交版模型：定稿(优化回流比) -> ASCII 文件名 .bkp/.apwz"""
import sys, time, os, shutil
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\_probe\rr_opt_s.bkp'
OUT_BKP = r'D:\<化工工作区>\NA-Chemical-10000t_Submit.bkp'
OUT_APWZ = r'D:\<化工工作区>\NA-Chemical-10000t_Submit.apwz'
L = []

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
while time.time() - t0 < 700:
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
for x in MSGS[-8:]:
    L.append(x[:140])

def g(p):
    n = doc.Tree.FindNode(p)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return None

# 关键结果复核
L.append('')
L.append('产品 S-209 = %s kg/h' % g(r'\Data\Streams\S-209\Output\MASSFLMX\MIXED'))
L.append('3-CP  S-121 = %s kg/h' % g(r'\Data\Streams\S-121\Output\MASSFLMX\MIXED'))
L.append('T-401 REB = %s Gcal/h ; T-404 REB = %s' % (
    g(r'\Data\Blocks\T-401\Output\REB_DUTY'), g(r'\Data\Blocks\T-404\Output\REB_DUTY')))
try:
    doc.SaveAs(OUT_BKP)
    L.append('BKP 已保存: %s' % OUT_BKP)
except Exception as ex:
    L.append('BKP 保存失败: %s' % ex)
try:
    doc.SaveAs2(OUT_APWZ)
    L.append('APWZ 已保存: %s' % OUT_APWZ)
except Exception as ex:
    L.append('APWZ(SaveAs2) 失败: %s' % ex)
try:
    doc.Close()
except Exception:
    pass

for p in [OUT_BKP, OUT_APWZ]:
    L.append('%s exists=%s size=%s' % (os.path.basename(p), os.path.exists(p), os.path.getsize(p) if os.path.exists(p) else 0))

open(r'D:\<化工工作区>\_probe\submit_model.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('\n'.join(L))
