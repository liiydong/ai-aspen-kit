# -*- coding: utf-8 -*-
"""测试亨利组分两种段名写法，看能否消除 T-202 的 O2 相平衡警告"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\_probe\wv4_nphase.bkp'   # 已含 NPHASE=3 的版本
OUT = r'D:\<化工工作区>\_probe'
base = open(SRC, encoding='utf-8', errors='ignore').read()

m3 = re.search(r'\?\s*COMPONENTS\s+"COMP-LIST"', base)
print('锚点:', bool(m3))
LST = '( N2 O2 CO CO2 HCN )'
V1 = '? COMPONENTS "HENRY-COMPS" MAIN ?\n\\ "HENRY-COMPS" CID = %s \\\n' % LST
V2 = '? COMPONENTS "HENRY-COMPS" ?\n\\ "HENRY-COMPS" CID = %s \\\n' % LST
VAR = {'H1': base[:m3.start()] + V1 + base[m3.start():],
       'H2': base[:m3.start()] + V2 + base[m3.start():]}


def run(src, tag):
    MSGS = []

    class Sink:
        def OnControlPanelMessage(self, *a):
            s = ' '.join(str(x) for x in a).strip()
            if s and s != 'False':
                MSGS.append(s)

    doc = win32.DispatchEx('Apwn.Document')
    doc.SuppressDialogs = True
    win32.WithEvents(doc, Sink)
    try:
        doc.InitFromArchive2(src)
        time.sleep(2)
        applied = []
        for p in [r'\Data\Components\Henry-Comps', r'\Data\Components\Henry-Comps\Input\CID']:
            try:
                nn = doc.Tree.FindNode(p)
                if nn is None:
                    applied.append('%s=None' % p.split('\\')[-1])
                else:
                    try:
                        applied.append('%s rows=%d' % (p.split('\\')[-1], nn.Elements.Count))
                    except Exception:
                        applied.append('%s leaf' % p.split('\\')[-1])
            except Exception:
                applied.append('ERR')
        doc.Engine.Run2(False)
        t1 = time.time()
        while time.time() - t1 < 700:
            pythoncom.PumpWaitingMessages()
            time.sleep(0.25)
            if time.time() - t1 > 10:
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
            try:
                return n.Value if n is not None else None
            except Exception:
                return None

        warn = [x for x in MSGS if re.search(r'\*\s*WARNING', x)]
        kinds = {}
        for w in warn:
            k = re.sub(r'\s+', ' ', w.replace('False', '').strip())[:64]
            kinds[k] = kinds.get(k, 0) + 1
        print('%-6s 警告=%-3d [%s]  产品=%.3f' % (tag, len(warn), '; '.join(applied),
                                              float(g(r'\Data\Streams\S-209\Output\MASSFLMX\MIXED') or 0)))
        for k, c in sorted(kinds.items(), key=lambda x: -x[1]):
            print('        ×%-3d %s' % (c, k))
        try:
            doc.Close()
        except Exception:
            pass
    except Exception as e:
        print('%-6s FAIL %s' % (tag, str(e)[:100]))


for tag, txt in VAR.items():
    p = os.path.join(OUT, 'hv_%s.bkp' % tag)
    open(p, 'w', encoding='utf-8', errors='ignore').write(txt)
    run(p, tag)
print('DONE')
