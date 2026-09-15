# -*- coding: utf-8 -*-
"""读取塔顶馏出物的液相密度与 T-402 进料组成，供水力学校核"""
import sys, time, re
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_BIP2.bkp'
doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
doc.InitFromArchive2(SRC)
time.sleep(2)


class S:
    def OnControlPanelMessage(self, *a):
        pass


win32.WithEvents(doc, S)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 400:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 8:
        try:
            if not doc.Engine.IsRunning:
                break
        except Exception:
            break
for _ in range(60):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)


def g(p):
    n = doc.Tree.FindNode(p)
    try:
        return n.Value if n is not None else None
    except Exception:
        return None


def comp(s):
    nn = doc.Tree.FindNode(r'\Data\Streams\%s\Output\MASSFLOW3' % s)
    d = {}
    if nn is None:
        return d
    for i in range(nn.Elements.Count):
        e = nn.Elements.Item(i)
        try:
            v = float(e.Value or 0)
            if abs(v) > 1e-6:
                d[e.Name] = v
        except Exception:
            pass
    return d


print('=== 液相密度 (kg/m3) ===')
for s, p in [('S-115', r'\Data\Streams\S-115\Output\RHOMX_MASS\MIXED'),
             ('S-117', r'\Data\Streams\S-117\Output\RHOMX_MASS\MIXED'),
             ('S-121', r'\Data\Streams\S-121\Output\RHOMX_MASS\MIXED'),
             ('S-116A', r'\Data\Streams\S-116A\Output\RHOMX_MASS\MIXED')]:
    print('  %-8s %s' % (s, g(p)))

print()
print('=== 温度 / 压力 ===')
for s in ['S-115', 'S-117', 'S-121', 'S-116A']:
    print('  %-8s T=%s P=%s' % (s, g(r'\Data\Streams\%s\Output\TEMP_OUT\MIXED' % s),
                                g(r'\Data\Streams\%s\Output\PRES_OUT\MIXED' % s)))

print()
print('=== 组成 (kg/h) ===')
for s in ['S-116A', 'S-117', 'S-115', 'S-121']:
    print('  %-8s %s' % (s, comp(s)))

print()
print('=== 塔顶压力剖面（第1板） ===')
for b in ['T-401', 'T-402', 'T-404']:
    print('  %-6s TOP_T=%s BOT_T=%s' % (b, g(r'\Data\Blocks\%s\Output\TOP_TEMP' % b),
                                        g(r'\Data\Blocks\%s\Output\BOTTOM_TEMP' % b)))
doc.Close()
print('DONE')
