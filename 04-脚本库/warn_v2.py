# -*- coding: utf-8 -*-
"""对照实验 v2：用 COM 写 FREE_WATER / NPHASE，并文本注入亨利组分"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_BIP2.bkp'
OUT = r'D:\<化工工作区>\_probe'
base = open(SRC, encoding='utf-8', errors='ignore').read()

m3 = re.search(r'\?\s*COMPONENTS\s+"COMP-LIST"', base)
if not m3:
    m3 = re.search(r'\?\s*COMPONENTS\s+MAIN\s*\?', base)
HENRY = ('? COMPONENTS "HENRY-COMPS" MAIN ?\n'
         '\\ "HENRY-COMPS" CID = ( N2 O2 CO CO2 HCN ) \\\n')
henry_txt = base[:m3.start()] + HENRY + base[m3.start():]
print('HENRY 锚点:', bool(m3))


def run(src, tag, sets=None):
    MSGS = []

    class Sink:
        def OnControlPanelMessage(self, *a):
            s = ' '.join(str(x) for x in a).strip()
            if s and s != 'False':
                MSGS.append(s)

    try:
        doc = win32.DispatchEx('Apwn.Document')
        doc.SuppressDialogs = True
        win32.WithEvents(doc, Sink)
        doc.InitFromArchive2(src)
        time.sleep(2)
        applied = []
        if sets:
            for path, val in sets:
                try:
                    n = doc.Tree.FindNode(path)
                    if n is None:
                        applied.append('%s=NOTFOUND' % path.split('\\')[-1]); continue
                    n.Value = val
                    applied.append('%s=%s' % (path.split('\\')[-1], n.Value))
                except Exception as e:
                    applied.append('%s ERR %s' % (path.split('\\')[-1], str(e)[:30]))
        doc.Engine.Run2(False)
        t1 = time.time()
        while time.time() - t1 < 600:
            pythoncom.PumpWaitingMessages()
            time.sleep(0.25)
            if time.time() - t1 > 10:
                try:
                    if not bool(doc.Engine.IsRunning):
                        break
                except Exception:
                    break
        for _ in range(70):
            pythoncom.PumpWaitingMessages()
            time.sleep(0.05)

        def g(p):
            n = doc.Tree.FindNode(p)
            try:
                return n.Value if n is not None else None
            except Exception:
                return None

        warn = [x for x in MSGS if re.search(r'\*\s*WARNING', x)]
        IN = ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-132', 'S-210']
        OUTL = ['S-130', 'S-131', 'S-211', 'S-212', 'S-213', 'S-115', 'S-119',
                'S-122', 'S-216', 'S-217', 'S-218', 'S-405C', 'S-406', 'S-209']
        vi = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in IN)
        vo = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in OUTL)
        prod = float(g(r'\Data\Streams\S-209\Output\MASSFLMX\MIXED') or 0)
        print('%-16s 警告=%-3d 设置[%s]  进=%.1f 出=%.1f 偏差=%.4f%%  产品=%.2f' % (
            tag, len(warn), '; '.join(applied), vi, vo,
            abs(vi - vo) / vi * 100 if vi else 0, prod))
        kinds = {}
        for w in warn:
            k = re.sub(r'\s+', ' ', w.replace('False', '').strip())[:60]
            kinds[k] = kinds.get(k, 0) + 1
        for k, c in sorted(kinds.items(), key=lambda x: -x[1]):
            print('        ×%-3d %s' % (c, k))
        try:
            doc.Close()
        except Exception:
            pass
    except Exception as e:
        print('%-16s FAIL %s' % (tag, str(e)[:110]))


def mk(txt, name):
    p = os.path.join(OUT, name)
    open(p, 'w', encoding='utf-8', errors='ignore').write(txt)
    return p


SET_FW = [(r'\Data\Setup\Main\Input\FREE_WATER', 'YES')]
SET_NP = [(r'\Data\Setup\Main\Input\NPHASE', 3)]

run(SRC, 'A_base')
run(mk(base, 'wv2_B.bkp'), 'B_freewater', SET_FW)
run(mk(base, 'wv2_C.bkp'), 'C_nphase3', SET_NP)
run(mk(base, 'wv2_D.bkp'), 'D_henry')
run(mk(henry_txt, 'wv2_E.bkp'), 'E_henry+FW', SET_FW)
print('DONE')
