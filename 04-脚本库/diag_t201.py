# -*- coding: utf-8 -*-
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_sim6.bkp'
T0 = open(BASE, encoding='utf-8', errors='ignore').read()

FEEDS_OLD = ('FEEDS FEED-SID = "S-107" FEED-STAGE = 1 FEED-CONVE2 = "ON-STAGE" / '
             'FEED-SID = "S-108" FEED-STAGE = 6 FEED-CONVE2 = "ON-STAGE"')


def run(tag, txt):
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
    print('=' * 22, tag, '面板', len(msgs), flush=True)
    for x in msgs[:22]:
        print('   |', x[:185], flush=True)
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
    for s in ['S-108', 'S-109', 'S-110']:
        nd = doc.Tree.FindNode(r'\Data\Streams\%s\Output' % s)
        if nd is not None:
            got = []
            for f in ['TEMP', 'PRES', 'MOLE-FLOW', 'MASS-FLOW', 'VFRAC']:
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


# A: 交换进料板
A = T0.replace(FEEDS_OLD,
               'FEEDS FEED-SID = "S-108" FEED-STAGE = 1 / FEED-SID = "S-107" FEED-STAGE = 6')
print('A 替换成功:', A != T0)
run('Aswap', A)
print()

# B: 提高进料温度到 200 ℃
m = re.search(r'(\? BLOCK HEATER "E-104" \? [^\n]*?\n?\\ PARAM TEMP = )100\.', T0)
if m:
    B = T0[:m.start(1)] + m.group(1) + '200.' + T0[m.end():]
    print('B 替换成功: True')
else:
    B = T0.replace('? BLOCK HEATER "E-104" ? ; "METCBAR_MOLE" ; ; HEATER ; \\ PARAM TEMP = 100.',
                   '? BLOCK HEATER "E-104" ? ; "METCBAR_MOLE" ; ; HEATER ; \\ PARAM TEMP = 200.')
    print('B 替换成功:', B != T0)
run('Bhot', B)
print()
print('DONE')
