# -*- coding: utf-8 -*-
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_sim6.bkp'
T0 = open(BASE, encoding='utf-8', errors='ignore').read()

pat = re.compile(r'(BLOCK\s+BLKID\s*=\s*"T-201"\s+BLKTYPE\s*=\s*"RADFRAC"\s+'
                 r'MDLTYPE\s*=\s*"RadFrac"\s+IN\s*=\s*\(\s*)([^)]*?)(\s*\))')
m = pat.search(T0)
print('匹配:', bool(m))
if m:
    print('原 IN =', re.sub(r'\s+', ' ', m.group(2)))

IN_NEW = {
    1: '"S-107" M0-1 "S-108" M1-2',
    2: '"S-108" M0-1 "S-107" M1-2',
}


def run(v):
    def rep(mm):
        return mm.group(1) + IN_NEW[v] + mm.group(3)
    txt = pat.sub(rep, T0, count=1)
    out = r'D:\<化工工作区>\_probe\port_%d.bkp' % v
    open(out, 'w', encoding='utf-8', errors='ignore').write(txt)
    # 校验
    m2 = pat.search(txt)
    print('=' * 25, '变体', v, '-> IN =', re.sub(r'\s+', ' ', m2.group(2)), flush=True)

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
    while time.time() - t1 < 260:
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

    print('   面板 %d 条' % len(msgs), flush=True)
    for x in msgs[:24]:
        print('   |', x[:190], flush=True)
    nn = doc.Tree.FindNode(r'\Data\Blocks\T-201\Output')
    if nn is not None:
        got = []
        for f in ['TEMP', 'PRES', 'DUTY', 'B-PRES']:
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
