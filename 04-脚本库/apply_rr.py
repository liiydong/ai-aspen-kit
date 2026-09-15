# -*- coding: utf-8 -*-
"""把回流比优化落到最终模型：T-401 3->2, T-402 15->8, T-404 10->4；跑通并采集水力学数据"""
import sys, re, time, os, json
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_BIP.bkp'
OUTF = r'D:\<化工工作区>\NA-Chemical-10000t_BIP2.bkp'
RR = {'T-401': ('3.0', '2.0'), 'T-402': ('15.0', '8.0'), 'T-404': ('10.0', '4.0')}

t = open(SRC, encoding='utf-8', errors='ignore').read()
for bid, (a, b) in RR.items():
    m = re.search(r'\?\s*BLOCK\s+RADFRAC\s+"?%s"?\s*\?' % re.escape(bid), t)
    m2 = re.search(r'\?\s*BLOCK\s', t[m.end():])
    end = m.end() + m2.start()
    seg = t[m.start():end]
    seg2, n = re.subn(r'BASIS-RR = %s <-1>' % re.escape(a), 'BASIS-RR = %s <-1>' % b, seg, count=1)
    if n == 0:  # 折行容错
        seg2, n = re.subn(r'BASIS-RR\s*=\s*%s\s*<-1>' % re.escape(a),
                          'BASIS-RR = %s <-1>' % b, seg, count=1)
    print('%s RR %s -> %s  替换=%d' % (bid, a, b, n))
    t = t[:m.start()] + seg2 + t[end:]
open(OUTF, 'w', encoding='utf-8', errors='ignore').write(t)
print('写出', OUTF, len(t))

MSGS = []


class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            MSGS.append(s)


doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
win32.WithEvents(doc, Sink)
doc.InitFromArchive2(OUTF)
time.sleep(2)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 600:
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


def g(p):
    n = doc.Tree.FindNode(p)
    try:
        return n.Value if n is not None else None
    except Exception:
        return None


print()
print('=== 错误计数 ===')
for x in MSGS:
    if re.search(r'Terminal Errors|Severe Errors|Errors\s+\d|Summary of Simulation', x):
        print('  ', x[:120])

IN = ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-132', 'S-210']
OUTL = ['S-130', 'S-131', 'S-211', 'S-212', 'S-213', 'S-115', 'S-119',
        'S-122', 'S-216', 'S-217', 'S-218', 'S-405C', 'S-406', 'S-209']
vi = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in IN)
vo = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in OUTL)
print('进=%.2f 出=%.2f 偏差=%.4f%%' % (vi, vo, abs(vi - vo) / vi * 100))

print()
print('%-7s %8s %8s %9s %9s %8s %9s %8s' % ('塔', 'TOP_T', 'BOT_T', 'TOP_LF', 'MOLE_D', 'BOT_VF', 'COND', 'REB'))
data = {}
for b in ['T-201', 'T-401', 'T-402', 'T-404']:
    d = dict(
        top_t=g(r'\Data\Blocks\%s\Output\TOP_TEMP' % b),
        bot_t=g(r'\Data\Blocks\%s\Output\BOTTOM_TEMP' % b),
        top_lf=g(r'\Data\Blocks\%s\Output\TOP_LFLOW' % b),
        mole_d=g(r'\Data\Blocks\%s\Output\MOLE_D' % b),
        bot_vf=g(r'\Data\Blocks\%s\Output\BOT_VFLOW' % b),
        cond=g(r'\Data\Blocks\%s\Output\COND_DUTY' % b),
        reb=g(r'\Data\Blocks\%s\Output\REB_DUTY' % b),
    )
    data[b] = d
    print('%-7s %8.2f %8.2f %9.3f %9.3f %8.3f %9.5f %8.5f' % (
        b, d['top_t'], d['bot_t'], d['top_lf'], d['mole_d'], d['bot_vf'], d['cond'], d['reb']))

print()
print('=== 产品 ===')
for s in ['S-209', 'S-115', 'S-117', 'S-121']:
    nn = doc.Tree.FindNode(r'\Data\Streams\%s\Output\MASSFLOW3' % s)
    items = []
    if nn is not None:
        for i in range(nn.Elements.Count):
            e = nn.Elements.Item(i)
            try:
                v = float(e.Value or 0)
                if abs(v) > 0.01:
                    items.append('%s=%.2f' % (e.Name, v))
            except Exception:
                pass
    print('  %-7s MASS=%-10s %s' % (s, g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s), ', '.join(items)))

doc.SaveAs(OUTF)
try:
    doc.Close()
except Exception:
    pass
print('DONE')
open(r'D:\<化工工作区>\_probe\rr_data.json', 'w', encoding='utf-8').write(json.dumps(data, ensure_ascii=False, indent=1))
