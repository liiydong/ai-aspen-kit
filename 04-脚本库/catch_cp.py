# -*- coding: utf-8 -*-
"""挂 ControlPanel 事件钩子 + Export(6) 抓 Aspen 报错"""
import os, time
import pythoncom
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\fix301.bkp'
OUT = r'D:\<化工工作区>\_probe\art2'
os.makedirs(OUT, exist_ok=True)
MSGS = []


class Sink:
    def OnControlPanelMessage(self, *a):
        MSGS.append(' '.join(str(x) for x in a))
        print('  [CP]', MSGS[-1][:300], flush=True)


doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
try:
    sink = win32.WithEvents(doc, Sink)
    print('事件钩子已挂上', flush=True)
except Exception as e:
    sink = None
    print('事件钩子挂载失败:', str(e)[:120], flush=True)

doc.InitFromArchive2(SRC)
time.sleep(4)
print('Ready =', doc.Engine.Ready, flush=True)

print('Run2(False) =', repr(doc.Engine.Run2(False)), flush=True)
t0 = time.time()
while time.time() - t0 < 180:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 12:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(30):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)
print('运行用时 %.1fs  捕获消息 %d 条' % (time.time() - t0, len(MSGS)), flush=True)

try:
    doc.SaveAs(os.path.join(OUT, 'CAP.bkp'))
except Exception as e:
    print('SaveAs 失败', str(e)[:80], flush=True)

for typ, fn in [(6, 'run.msg'), (2, 'run.rep'), (3, 'run.sum')]:
    try:
        doc.Export(typ, os.path.join(OUT, fn))
    except Exception as e:
        print('Export(%d) %s' % (typ, str(e)[:70]), flush=True)

print()
print('=== 目录 ===', flush=True)
for f in sorted(os.listdir(OUT)):
    print('   %-26s %d' % (f, os.path.getsize(os.path.join(OUT, f))), flush=True)

print()
print('=== 捕获到的控制面板消息 ===', flush=True)
for m in MSGS[:80]:
    print('   ', m[:400], flush=True)
if not MSGS:
    print('   （没有事件消息）', flush=True)

for fn in os.listdir(OUT):
    if fn.lower().endswith(('.cpm', '.msg', '.rep', '.sum')):
        p = os.path.join(OUT, fn)
        txt = open(p, encoding='utf-8', errors='ignore').read()
        print()
        print('=' * 20, fn, '(%d 字)' % len(txt), flush=True)
        print(txt[:6000], flush=True)
try:
    doc.Close()
except Exception:
    pass
