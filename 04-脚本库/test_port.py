# -*- coding: utf-8 -*-
"""T-201 进料端口分离测试"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_sim6.bkp'
T0 = open(BASE, encoding='utf-8', errors='ignore').read()

OLD = 'BLOCK BLKID = "T-201" BLKTYPE = "RADFRAC" MDLTYPE = "RadFrac" IN = ( "S-107" M0-1 "S-108" M0-1 )'
print('原连接存在:', OLD in T0)

VARIANTS = {
    1: ('BLOCK BLKID = "T-201" BLKTYPE = "RADFRAC" MDLTYPE = "RadFrac" '
        'IN = ( "S-107" M0-1 "S-108" M1-2 )'),
    2: ('BLOCK BLKID = "T-201" BLKTYPE = "RADFRAC" MDLTYPE = "RadFrac" '
        'IN = ( "S-108" M0-1 "S-107" M1-2 )'),
}


def run(v):
    txt = T0.replace(OLD, VARIANTS[v])
    out = r'D:\<化工工作区>\_probe\port_%d.bkp' % v
    open(out, 'w', encoding='utf-8', errors='ignore').write(txt)
    msgs = []

    class Sink:
        def OnControlPanelMessage(self, *a):
            s = ' '.join(str(x) for x in a).strip()
            if s and s != 'False':
                msgs.append(s)

    doc = win32.DispatchEx('Apwn.Document')
    try:
        doc.SuppressDialogs = True
    except Exception:
        pass
    try:
        win32.WithEvents(doc, Sink)
    except Exception:
        pass
    doc.InitFromArchive2(out)
    time.sleep(3)
    doc.Engine.Run2(False)
    t1 = time.time()
    while time.time() - t1 < 240:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.25)
        if time.time() - t1 > 8:
            try:
                if not bool(doc.Engine.IsRunning):
                    break
            except Exception:
                break
    for _ in range(40):
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)

    print('=' * 25, '端口变体', v, '面板', len(msgs), '条', flush=True)
    bad = [x for x in msgs if 'ERROR' in x or 'SEVERE' in x or 'T-201' in x]
    for x in bad[:14]:
        print('   |', x[:200], flush=True)
    if not bad:
        print('   ✓ 无错误')
    nn = doc.Tree.FindNode(r'\Data\Blocks\T-201\Output')
    if nn is not None:
        got = []
        for f in ['TEMP', 'PRES', 'DUTY', 'B-PRES', 'COND_DUTY', 'REB_DUTY']:
            try:
                e = nn.Elements.Item(f)
                if e is not None and e.Value not in (None, ''):
                    got.append('%s=%s' % (f, e.Value))
            except Exception:
                pass
        print('   T-201:', ', '.join(got) if got else '(空)', flush=True)
    for s in ['S-109', 'S-110', 'S-112']:
        nd = doc.Tree.FindNode(r'\Data\Streams\%s\Output' % s)
        if nd is not None:
            got = []
            for f in ['TEMP', 'PRES', 'MOLE-FLOW', 'MASS-FLOW']:
                try:
                    e = nd.Elements.Item(f)
                    if e is not None and e.Value not in (None, ''):
                        got.append('%s=%s' % (f, e.Value))
                except Exception:
                    pass
            print('   %s: %s' % (s, ', '.join(got) if got else '(空)'), flush=True)
    try:
        doc.SaveAs(r'D:\<化工工作区>\_probe\port_%d_saved.bkp' % v)
    except Exception:
        pass
    try:
        doc.Close()
    except Exception:
        pass
    time.sleep(1)


for v in (1, 2):
    try:
        run(v)
    except Exception as ex:
        print('变体 %d 异常 %s' % (v, ex))
    print()
print('DONE')
