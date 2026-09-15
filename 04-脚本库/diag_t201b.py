# -*- coding: utf-8 -*-
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_sim6.bkp'
T0 = open(BASE, encoding='utf-8', errors='ignore').read()

# 定位并打印 T-201 的 FEEDS 与 E-104
fp = re.compile(r'FEEDS\s+FEED-SID\s*=\s*"S-107"\s+FEED-STAGE\s*=\s*(\d+)(.{0,40}?)'
                r'FEED-SID\s*=\s*"S-108"\s+FEED-STAGE\s*=\s*(\d+)', re.S)
m = fp.search(T0)
print('FEEDS 匹配:', bool(m), '->', repr(m.group(0)[:120]) if m else '')
ep = re.compile(r'BLOCK HEATER "E-104"[\s\S]{0,300}?TEMP\s*=\s*([\d.]+)')
me = ep.search(T0)
print('E-104 匹配:', bool(me), '->', repr(me.group(0)[-60:]) if me else '')


def load_run(tag, txt):
    out = r'D:\<化工工作区>\_probe\dg_%s.bkp' % tag
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
    print('=' * 20, tag, '面板', len(msgs), flush=True)
    for x in msgs[:26]:
        print('   |', x[:180], flush=True)
    nn = doc.Tree.FindNode(r'\Data\Blocks\T-201\Output')
    if nn is not None:
        got = []
        for f in ['TEMP', 'PRES', 'DUTY', 'B-PRES']:
            try:
                e = nn.Elements.Item(f)
                if e is not None:
                    got.append('%s=%s' % (f, e.Value))
            except Exception:
                pass
        print('   T-201:', ', '.join(got), flush=True)
    for s in ['S-108', 'S-110']:
        nd = doc.Tree.FindNode(r'\Data\Streams\%s\Output' % s)
        if nd is not None:
            got = []
            for f in ['TEMP', 'PRES', 'MOLE-FLOW', 'VFRAC']:
                try:
                    e = nd.Elements.Item(f)
                    if e is not None and e.Value not in (None, ''):
                        got.append('%s=%s' % (f, e.Value))
                except Exception:
                    pass
            print('   %s: %s' % (s, ', '.join(got) if got else '(空)'), flush=True)
    try:
        doc.SaveAs(r'D:\<化工工作区>\_probe\dg_%s_saved.bkp' % tag)
    except Exception:
        pass
    try:
        doc.Close()
    except Exception:
        pass
    time.sleep(1)


# A: 交换进料板 1<->6
A = fp.sub(lambda mm: mm.group(0).replace('FEED-STAGE = %s' % mm.group(1), 'FEED-STAGE = @1@')
           .replace('FEED-STAGE = %s' % mm.group(2), 'FEED-STAGE = @2@')
           .replace('@1@', mm.group(2)).replace('@2@', mm.group(1)), T0, count=1)
print('A 是否变化:', A != T0)
load_run('Aswap', A)
print()

# B: E-104 提到 250 ℃
if me:
    B = T0[:me.start(1)] + '250.' + T0[me.end(1):]
    print('B 是否变化:', B != T0)
    load_run('Bhot', B)
print()
print('DONE')
