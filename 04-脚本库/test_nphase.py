# -*- coding: utf-8 -*-
"""测试：SETUP SIM-OPTIONS 里开三相闪蒸（NPHASE=3）"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_BIP2.bkp'
OUT = r'D:\<化工工作区>\_probe\wv4_nphase.bkp'
base = open(SRC, encoding='utf-8', errors='ignore').read()

pat = re.compile(r'(\?\s*SETUP\s+"SIM-OPTIONS"\s*\?\s*;\s*"[A-Z_]+"\s*;\s*)\?')
m = pat.search(base)
print('SIM-OPTIONS 锚点:', bool(m))
print('  原文:', repr(base[m.start():m.end()]) if m else '')
REPL = r'\1\\ "SIM-OPTIONS" NPHASE = 3 NPHASE-HI = 3 NPHASE-RB = 3 \\ \ ?'
txt, n = pat.subn(REPL, base, count=1)
print('替换次数:', n)
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
for p in [r'\Data\Setup\Sim-Options\Input\NPHASE', r'\Data\Setup\Main\Input\NPHASE']:
    try:
        nn = doc.Tree.FindNode(p)
        print('  %-42s = %s' % (p, (nn.Value if nn is not None else None)))
    except Exception as e:
        print('  %-42s ERR' % p)
sys.stdout.flush()

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
print()
print('警告数 =', len(warn))
kinds = {}
for w in warn:
    k = re.sub(r'\s+', ' ', w.replace('False', '').strip())[:70]
    kinds[k] = kinds.get(k, 0) + 1
for k, c in sorted(kinds.items(), key=lambda x: -x[1]):
    print('   ×%-3d %s' % (c, k))
for x in MSGS:
    if re.search(r'Terminal Errors|Severe Errors|Errors\s+\d|Warnings\s+\d', x):
        print('  ', x.replace('False', '').strip()[:110])
IN = ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-132', 'S-210']
OUTL = ['S-130', 'S-131', 'S-211', 'S-212', 'S-213', 'S-115', 'S-119',
        'S-122', 'S-216', 'S-217', 'S-218', 'S-405C', 'S-406', 'S-209']
vi = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in IN)
vo = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in OUTL)
print('进=%.2f 出=%.2f 偏差=%.4f%%  产品=%.3f  S-113=%.1f' % (
    vi, vo, abs(vi - vo) / vi * 100 if vi else 0,
    float(g(r'\Data\Streams\S-209\Output\MASSFLMX\MIXED') or 0),
    float(g(r'\Data\Streams\S-113\Output\MASSFLMX\MIXED') or 0)))
print('DONE')
