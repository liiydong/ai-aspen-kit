# -*- coding: utf-8 -*-
"""v9-A：补 DATABANKS 清单 + 改回 NRTL，看 NaOH 缺参数是否消失、3-CP 是否被吸收"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_v7.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_v9a.bkp'

t = open(SRC, encoding='utf-8', errors='ignore').read()

print('=== 当前 DATABANKS 段 ===', flush=True)
m = re.search(r'\?\s*DATABANKS\s*\?[^?]{0,400}', t)
print(repr(t[m.start():m.start() + 400]) if m else '未找到', flush=True)

m2 = re.search(r'\?\s*DATABANKS\s*\?', t)
if m2:
    end = t.find('\n?', m2.end())
    if end < 0:
        end = m2.end() + 400
    nxt = re.search(r'\\ \?', t[m2.end():])
    ins = m2.end()
    DB = (' \\ FILE-SYM-NAM = ( "APV150 PURE41" "APV150 AQUEOUS" "APV150 SOLIDS" '
          '"APV150 INORGANIC" "APESV150 AP-EOS" "NISTV150 NIST-TRC" ) ')
    t = t[:ins] + DB + t[ins:]
    print('已插入 DATABANKS 清单', flush=True)

m3 = re.search(r'PARAM\s+BASE\s*=\s*"[^"]*"', t)
print('当前 BASE:', repr(t[m3.start():m3.end()]) if m3 else '?', flush=True)
if m3:
    t = t[:m3.start()] + 'PARAM BASE = "NRTL"' + t[m3.end():]
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)
print('写出:', OUT, flush=True)

MSGS = []


class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            MSGS.append(s)


doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
try:
    win32.WithEvents(doc, Sink)
except Exception:
    pass
doc.InitFromArchive2(OUT)
time.sleep(3)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 420:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 10:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(60):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)

print()
print('=== 面板（全部，最多 60 条） ===', flush=True)
for i, x in enumerate(MSGS[:60], 1):
    print('%3d| %s' % (i, x[:175]), flush=True)


def gg(p):
    n = doc.Tree.FindNode(p)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return None


print()
print('=== S-110 / S-109 组成（看 3-CP 是否被吸收） ===', flush=True)
for s in ['S-109', 'S-110', 'S-114', 'S-116']:
    b = r'\Data\Streams\%s\Output' % s
    print('  --- %s  MASS=%s  MOLE=%s  VFRAC=%s' % (
        s, gg(b + r'\MASSFLMX\MIXED'), gg(b + r'\MOLEFLMX\MIXED'),
        gg(b + r'\VFRAC_OUT\MIXED')), flush=True)
    n = doc.Tree.FindNode(b + r'\MASSFLOW3')
    if n is not None:
        for i in range(n.Elements.Count):
            try:
                e = n.Elements.Item(i)
                if e.Value and abs(float(e.Value)) > 0.01:
                    print('        %-8s %9.2f' % (e.Name, float(e.Value)), flush=True)
            except Exception:
                pass

try:
    doc.SaveAs(OUT.replace('.bkp', '.apwz'))
except Exception:
    pass
try:
    doc.Close()
except Exception:
    pass
print('DONE', flush=True)
