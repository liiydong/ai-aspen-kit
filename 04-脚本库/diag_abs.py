# -*- coding: utf-8 -*-
"""诊断 T-201 吸收塔：读 S-107/108/109/110 的组成与相态"""
import sys, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_v7.bkp'
MSGS = []


class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            MSGS.append(s)


doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
try:
    win32.WithEvents(doc, Sink)
except Exception:
    pass
doc.InitFromArchive2(SRC)
time.sleep(3)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 420:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 10:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(60):
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


def show(s):
    b = r'\Data\Streams\%s\Output' % s
    print('=== %s  T=%s  P=%s  MOLE=%.4f kmol/h  MASS=%.1f kg/h  VFRAC=%s' % (
        s, g(b + r'\TEMP_OUT'), g(b + r'\PRES_OUT'),
        g(b + r'\MOLEFLMX\MIXED') or 0, g(b + r'\MASSFLMX\MIXED') or 0,
        g(b + r'\VFRAC_OUT\MIXED')), flush=True)
    n = doc.Tree.FindNode(b + r'\MASSFLOW3')
    if n is None:
        print('   (无组分数据)'); return
    for i in range(n.Elements.Count):
        try:
            e = n.Elements.Item(i)
            v = e.Value
            if v and abs(float(v)) > 1e-4:
                print('      %-8s %9.2f kg/h' % (e.Name, float(v)), flush=True)
        except Exception:
            pass


for s in ['S-106', 'S-108', 'S-109', 'S-107', 'S-110']:
    show(s)
    print()

print('=== T-201 输出节点 ===', flush=True)
n = doc.Tree.FindNode(r'\Data\Blocks\T-201\Output')
if n is not None:
    for i in range(n.Elements.Count):
        try:
            e = n.Elements.Item(i)
            v = None
            try:
                v = e.Value
            except Exception:
                pass
            if v not in (None, '', 0, 0.0):
                print('   %-16s %s' % (e.Name, v), flush=True)
        except Exception:
            pass

print()
print('=== 面板关键行 ===', flush=True)
for x in MSGS:
    if 'ERROR' in x or 'WARNING' in x or 'Converged' in x or 'completed' in x:
        print('   |', x[:180], flush=True)

try:
    doc.Close()
except Exception:
    pass
print('DONE')
