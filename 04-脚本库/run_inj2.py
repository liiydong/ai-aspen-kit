# -*- coding: utf-8 -*-
"""用正确的 COM 等待方式跑注入参数后的模型，并抓控制面板"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\_probe\inj.bkp'
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
print('\Data children =', doc.Tree.FindNode(r'\Data').Elements.Count)
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
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return None


print()
print('=== 控制面板关键消息 ===')
for i, x in enumerate(MSGS):
    if re.search(r'SEVERE|Terminal|Errors|Warnings|completed|BINARY|PARAM|UNIFAC', x, re.I):
        print('%3d | %s' % (i, x[:170]))

IN = ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-132', 'S-210']
OUTL = ['S-130', 'S-131', 'S-211', 'S-212', 'S-213', 'S-115', 'S-119',
        'S-122', 'S-216', 'S-217', 'S-218', 'S-405C', 'S-406', 'S-209']
vi = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in IN)
vo = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in OUTL)
print()
print('进料合计 = %.2f  出料合计 = %.2f  偏差 = %.4f%%' % (vi, vo, abs(vi - vo) / vi * 100 if vi else 0))

STREAMS = ['S-106', 'S-114', 'S-117', 'S-121', 'S-209', 'S-403', 'S-405B']
print()
print('%-8s %14s %10s' % ('流股', 'MASS kg/h', 'T C'))
for s in STREAMS:
    print('%-8s %14s %10s' % (s, g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s),
                              g(r'\Data\Streams\%s\Output\TEMP_OUT\MIXED' % s)))

# 产品组成
nn = doc.Tree.FindNode(r'\Data\Streams\S-209\Output\MASSFLOW3')
if nn is not None:
    print()
    print('--- S-209 组成 ---')
    for i in range(nn.Elements.Count):
        e = nn.Elements.Item(i)
        try:
            v = float(e.Value or 0)
            if abs(v) > 0.01:
                print('   %-14s %.4f kg/h' % (e.Name, v))
        except Exception:
            pass

# 关键塔热负荷
for b in ['T-401', 'T-402', 'T-403', 'T-404', 'T-201']:
    q = g(r'\Data\Blocks\%s\Output\COND_DUTY' % b)
    q2 = g(r'\Data\Blocks\%s\Output\REB_DUTY' % b)
    print('%-8s COND=%-12s REB=%s' % (b, q, q2))

doc.SaveAs(r'D:\<化工工作区>\NA_withBIP.bkp')
try:
    doc.Close()
except Exception:
    pass
try:
    doc.Quit()
except Exception:
    pass
print('DONE')
