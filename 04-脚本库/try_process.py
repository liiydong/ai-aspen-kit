# -*- coding: utf-8 -*-
"""试探 Engine.ProcessInput / ReinitializeEO / SynchronizeEO / Activate"""
import time, inspect
import win32com.client as win32

F = r'D:\<化工工作区>\_probe\full3.bkp'

CASES = [
    ('Engine.ProcessInput()', [lambda d: d.Engine.ProcessInput()]),
    ('doc.Activate() + Engine.Run2()', [lambda d: d.Activate(), lambda d: d.Engine.Run2()]),
    ('Engine.ReinitializeEO + Run2', [lambda d: d.Engine.ReinitializeEO(),
                                      lambda d: d.Engine.Run2()]),
    ('Engine.ProcessInput + Run2', [lambda d: d.Engine.ProcessInput(),
                                    lambda d: d.Engine.Run2()]),
    ('Engine.Runs + SynchronizeEO', [lambda d: d.Engine.Run2()]),
]

for tag, fns in CASES:
    print('=' * 18, tag, flush=True)
    doc = win32.DispatchEx('Apwn.Document')
    try:
        doc.SuppressDialogs = True
    except Exception:
        pass
    try:
        doc.InitFromArchive(F)
    except Exception as e:
        print('  打开失败', str(e)[:80], flush=True)
        continue
    time.sleep(4)
    ok = True
    for fn in fns:
        try:
            fn(doc)
        except Exception as e:
            print('  调用失败:', str(e)[:100], flush=True)
            ok = False
    if not ok:
        try:
            doc.Close()
        except Exception:
            pass
        print(flush=True)
        continue
    t0 = time.time()
    while time.time() - t0 < 40:
        time.sleep(3)
        try:
            if doc.Engine.IsRunning is False:
                break
        except Exception:
            break
    nn = doc.Tree.FindNode(r'\Data\Blocks\T-404\Output\TEMP')
    m = doc.Tree.FindNode(r'\Data\Blocks\M-101\Output\TEMP')
    print('   Ready=%r  T-404.TEMP=%r  M-101.TEMP=%r' % (
        doc.Engine.Ready, nn.Value if nn is not None else None,
        m.Value if m is not None else None), flush=True)
    try:
        doc.Close()
    except Exception:
        pass
    print(flush=True)
