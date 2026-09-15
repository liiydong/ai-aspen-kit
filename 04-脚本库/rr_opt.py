# -*- coding: utf-8 -*-
"""回流比优化：T-401 3->2, T-402 15->8, T-404 10->4"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_最终定稿.bkp'
OUT = r'D:\<化工工作区>\_probe\rr_opt.bkp'
t = open(SRC, encoding='utf-8', errors='ignore').read()
L = []

def find_block_span(x, bid):
    starts = [s.start() for s in re.finditer(r'\?\s*BLOCK\s+', x)]
    for k in range(len(starts)):
        s0 = starts[k]
        s1 = starts[k + 1] if k + 1 < len(starts) else len(x)
        m = re.match(r'\?\s*BLOCK\s*\n?\s*([A-Z0-9]+)\s*\n?\s*"([A-Za-z0-9-]+)"', x[s0:s1])
        if m and m.group(2) == bid:
            return s0, s1
    raise RuntimeError(bid)

def set_rr(x, bid, old, new):
    s0, s1 = find_block_span(x, bid)
    seg = x[s0:s1]
    seg2, n = re.subn(r'BASIS-RR\s*=\s*%s' % re.escape(old), 'BASIS-RR = %s' % new, seg)
    return x[:s0] + seg2 + x[s1:], n

for bid, old, new in [('T-401', '3.0', '2.0'), ('T-402', '15.0', '8.0'), ('T-404', '10.0', '4.0')]:
    t, n = set_rr(t, bid, old, new)
    L.append('%s RR %s -> %s  (%d 处)' % (bid, old, new, n))

open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)

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
time.sleep(3)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 700:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 10:
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
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return None

def comp(s):
    nn = doc.Tree.FindNode(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s)
    if nn is None:
        return None
    try:
        return float(nn.Value)
    except Exception:
        return None

def compw(s):
    nn = doc.Tree.FindNode(r'\Data\Streams\%s\Output\MASSFLOW3' % s)
    d = {}
    if nn is None:
        return d
    try:
        for i in range(nn.Elements.Count):
            e = nn.Elements.Item(i)
            try:
                v = float(e.Value) if e.Value is not None else 0.0
                if abs(v) > 1e-6:
                    d[e.Name] = v
            except Exception:
                pass
    except Exception:
        pass
    return d

L.append('')
L.append('===== 面板尾 =====')
for x in MSGS[-10:]:
    L.append('  ' + x[:150])
L.append('')
L.append('===== 结果 =====')
for b in ['T-401', 'T-402', 'T-404']:
    L.append('  %-7s RR=%s COND=%s Gcal/h REB=%s Gcal/h TOP=%s K BOT=%s K' % (
        b, g(r'\Data\Blocks\%s\Output\RR' % b),
        g(r'\Data\Blocks\%s\Output\COND_DUTY' % b),
        g(r'\Data\Blocks\%s\Output\REB_DUTY' % b),
        g(r'\Data\Blocks\%s\Output\TOP_TEMP' % b),
        g(r'\Data\Blocks\%s\Output\BOTTOM_TEMP' % b)))

IN = ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-132', 'S-210']
OUTL = ['S-130', 'S-131', 'S-211', 'S-212', 'S-213', 'S-115', 'S-119', 'S-122',
        'S-216', 'S-217', 'S-218', 'S-405C', 'S-406', 'S-209']
vi = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in IN)
vo = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in OUTL)
L.append('  物料平衡: 进 %.2f 出 %.2f 偏差 %.5f%%' % (vi, vo, (vo - vi) / vi * 100))
cs = ['S-209', 'S-121', 'S-115', 'S-114', 'S-119', 'S-117', 'S-116', 'S-118', 'S-120']
for s in cs:
    c = compw(s)
    tt = sum(c.values())
    L.append('  %-7s %8.2f kg/h : %s' % (s, tt, ', '.join('%s %.2f%%' % (k, v / tt * 100) for k, v in sorted(c.items(), key=lambda x: -x[1])[:4]) if tt else ''))

p = compw('S-209'); tt = sum(p.values())
nm = p.get('NAM', 0); mp3 = compw('S-101').get('3-MP', 0)
L.append('  产品 %.2f kg/h  折年产 %.0f t/a  NAM %.3f%%  3-MP 单耗 %.1f kg/t' % (
    tt, tt / 1000 * 7200, nm / tt * 100 if tt else 0, mp3 / nm * 1000 if nm else 0))
open(r'D:\<化工工作区>\_probe\rr_opt.txt', 'w', encoding='utf-8').write('\n'.join(L))
try:
    doc.SaveAs(r'D:\<化工工作区>\_probe\rr_opt_s.bkp')
except Exception as ex:
    L.append('save fail %s' % ex)
try:
    doc.Close()
except Exception:
    pass
print('DONE')
