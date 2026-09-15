# -*- coding: utf-8 -*-
"""单独测试：文本注入亨利组分（N2 O2 CO CO2 HCN），核对警告变化"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_BIP2.bkp'
OUT = r'D:\<化工工作区>\_probe\wv3_D_henry.bkp'
base = open(SRC, encoding='utf-8', errors='ignore').read()

m3 = re.search(r'\?\s*COMPONENTS\s+"COMP-LIST"', base)
print('COMP-LIST 锚点:', bool(m3))
if not m3:
    m3 = re.search(r'\?\s*COMPONENTS\s+MAIN\s*\?', base)
    print('  (改用 MAIN 段起点作锚点)')
HENRY = '? COMPONENTS "HENRY-COMPS" MAIN ?\n\\ "HENRY-COMPS" CID = ( N2 O2 CO CO2 HCN ) \\\n'
txt = base[:m3.start()] + HENRY + base[m3.start():]
open(OUT, 'w', encoding='utf-8', errors='ignore').write(txt)
print('写出', OUT, len(txt))
sys.stdout.flush()

MSGS = []


class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            MSGS.append(s)


doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
win32.WithEvents(doc, Sink)
doc.InitFromArchive2(OUT)
time.sleep(2)

# 验证亨利组分是否被接受
for p in [r'\Data\Components\Henry-Comps\Input\CID',
          r'\Data\Components\Henry-Comps\Input',
          r'\Data\Components\Henry-Comps']:
    try:
        n = doc.Tree.FindNode(p)
        if n is None:
            print('  %-46s None' % p)
        else:
            try:
                print('  %-46s rows=%d' % (p, n.Elements.Count))
            except Exception:
                print('  %-46s leaf' % p)
    except Exception as e:
        print('  %-46s ERR %s' % (p, str(e)[:40]))
sys.stdout.flush()

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
print()
print('警告数 =', len(warn))
kinds = {}
for w in warn:
    k = re.sub(r'\s+', ' ', w.replace('False', '').strip())[:70]
    kinds[k] = kinds.get(k, 0) + 1
for k, c in sorted(kinds.items(), key=lambda x: -x[1]):
    print('   ×%-3d %s' % (c, k))
IN = ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-132', 'S-210']
OUTL = ['S-130', 'S-131', 'S-211', 'S-212', 'S-213', 'S-115', 'S-119',
        'S-122', 'S-216', 'S-217', 'S-218', 'S-405C', 'S-406', 'S-209']
vi = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in IN)
vo = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in OUTL)
print('进=%.2f 出=%.2f 偏差=%.4f%%  产品=%.3f' % (vi, vo, abs(vi - vo) / vi * 100 if vi else 0,
                                            float(g(r'\Data\Streams\S-209\Output\MASSFLMX\MIXED') or 0)))
print('DONE')
