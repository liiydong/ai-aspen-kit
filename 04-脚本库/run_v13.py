# -*- coding: utf-8 -*-
"""v13：标定 T-402/403/404 采出量（按已算出的进料组成）并复算"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_v12.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_v13.bkp'

NEWDF = {'T-402': 0.80, 'T-403': 0.010, 'T-404': 0.950, 'T-401': 0.50}

t = open(SRC, encoding='utf-8', errors='ignore').read()
starts = [m.start() for m in re.finditer(r'\?\s*BLOCK\s+', t)]
for i in range(len(starts) - 1, -1, -1):
    s0 = starts[i]
    s1 = starts[i + 1] if i + 1 < len(starts) else len(t)
    m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
    if not m:
        continue
    bid = m.group(2)
    if bid in NEWDF:
        seg = t[s0:s1]
        v = NEWDF[bid]
        seg2 = re.sub(r'D:F = [\d.]+ <-1> <0>', 'D:F = %.4f <-1> <0>' % v, seg, count=1)
        if seg2 != seg:
            t = t[:s0] + seg2 + t[s1:]
            print('%s D:F -> %.4f' % (bid, v), flush=True)
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
while time.time() - t0 < 520:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 10:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(70):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)

print()
print('=== 面板尾部 ===', flush=True)
for x in MSGS[-24:]:
    print('   |', x[:175], flush=True)


def gg(p):
    n = doc.Tree.FindNode(p)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return None


print()
print('=== 精制段流股 ===', flush=True)
for s in ['S-116', 'S-117', 'S-118', 'S-119', 'S-120', 'S-121', 'S-122']:
    b = r'\Data\Streams\%s\Output' % s
    print('  --- %s  MASS=%s  T=%s' % (
        s, gg(b + r'\MASSFLMX\MIXED'), gg(b + r'\TEMP_OUT')), flush=True)
    n = doc.Tree.FindNode(b + r'\MASSFLOW3')
    if n is not None:
        for i in range(n.Elements.Count):
            try:
                e = n.Elements.Item(i)
                if e.Value and abs(float(e.Value)) > 0.05:
                    print('        %-8s %9.2f' % (e.Name, float(e.Value)), flush=True)
            except Exception:
                pass

print()
print('=== 塔结果 ===', flush=True)
for b in ['T-201', 'T-301', 'T-401', 'T-402', 'T-403', 'T-404']:
    bb = r'\Data\Blocks\%s\Output' % b
    print('  %-7s COND=%s REB=%s RR=%s Ttop=%s Tbot=%s' % (
        b, gg(bb + r'\COND_DUTY'), gg(bb + r'\REB_DUTY'), gg(bb + r'\MOLE_RR'),
        gg(bb + r'\TOP_TEMP'), gg(bb + r'\BOTTOM_TEMP')), flush=True)

for p in [OUT.replace('.bkp', '_s.bkp'), OUT.replace('.bkp', '.apwz')]:
    try:
        doc.SaveAs(p)
        print('  已保存:', p, os.path.getsize(p), flush=True)
    except Exception as ex:
        print('  保存失败:', ex, flush=True)
try:
    doc.Close()
except Exception:
    pass
print('DONE', flush=True)
