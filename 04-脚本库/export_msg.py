# -*- coding: utf-8 -*-
"""用 Export(6/2/3) 导出控制面板消息，拿到报错"""
import os, time
import pythoncom
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\fix301.bkp'
OUT = r'D:\<化工工作区>\_probe\art'
os.makedirs(OUT, exist_ok=True)

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive2(SRC)
time.sleep(4)
print('Ready =', doc.Engine.Ready, flush=True)

try:
    doc.Engine.ProcessInput()
    print('ProcessInput 返回', flush=True)
except Exception as e:
    print('ProcessInput:', str(e)[:90], flush=True)

print('Run2(False) =', repr(doc.Engine.Run2(False)), flush=True)
t0 = time.time()
while time.time() - t0 < 200:
    pythoncom.PumpWaitingMessages()
    try:
        if not bool(doc.Engine.IsRunning):
            break
    except Exception:
        break
    time.sleep(0.25)
print('运行用时 %.1fs' % (time.time() - t0), flush=True)

capture = os.path.join(OUT, 'RUN_CAPTURE.bkp')
try:
    doc.SaveAs(capture)
    print('SaveAs ok', flush=True)
except Exception as e:
    print('SaveAs 失败', str(e)[:90], flush=True)

for nm, typ, fn in (('messages', 6, 'run.msg'), ('report', 2, 'run.rep'),
                    ('summary', 3, 'run.sum'), ('input', 1, 'run.inp')):
    p = os.path.join(OUT, fn)
    try:
        doc.Export(typ, p)
        print('Export(%d) -> %s  %d bytes' % (typ, fn, os.path.getsize(p)), flush=True)
    except Exception as e:
        print('Export(%d) 失败 %s' % (typ, str(e)[:80]), flush=True)

print()
print('=== 输出目录 ===', flush=True)
for f in sorted(os.listdir(OUT)):
    print('   %-24s %d' % (f, os.path.getsize(os.path.join(OUT, f))), flush=True)

print()
for fn in ['run.msg', 'run.rep', 'run.sum']:
    p = os.path.join(OUT, fn)
    if not os.path.exists(p):
        continue
    txt = open(p, encoding='utf-8', errors='ignore').read()
    print('=' * 22, fn, '(%d 字)' % len(txt), flush=True)
    print(txt[:5000], flush=True)
    print(flush=True)
try:
    doc.Close()
except Exception:
    pass
