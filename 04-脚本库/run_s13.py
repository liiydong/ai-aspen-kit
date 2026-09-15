# -*- coding: utf-8 -*-
"""S13：母液排放比 0.12->0.08（精制收率锁 97%）+ 采齐论文所需组成"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_s12.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_s13.bkp'
t = open(SRC, encoding='utf-8', errors='ignore').read()
print('源大小', len(t))


def find_block_span(x, bid):
    starts = [s.start() for s in re.finditer(r'\?\s*BLOCK\s+', x)]
    for k in range(len(starts)):
        s0 = starts[k]
        s1 = starts[k + 1] if k + 1 < len(starts) else len(x)
        m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', x[s0:s1])
        if m and m.group(2) == bid:
            return s0, s1
    raise RuntimeError(bid)


def repl_block(x, bid, fn):
    s0, s1 = find_block_span(x, bid)
    return x[:s0] + fn(x[s0:s1]) + x[s1:]


t = repl_block(t, 'M-503', lambda s: re.sub(r'(FRACS\s*=\s*)0\.12(\s*<0>)', r'\g<1>0.08\g<2>', s))
i = t.find('BLOCK SEP "M-503"')
print('1) M-503 排放比 0.12 -> 0.08  %s' % re.findall(r'FRACS = [\d.]+', t[i:i + 3000])[:3])

open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)
print('写出:', OUT, os.path.getsize(OUT))

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
while time.time() - t0 < 600:
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
    nn = doc.Tree.FindNode(r'\Data\Streams\%s\Output\MASSFLOW3' % s)
    d = {}
    if nn is None:
        return d
    try:
        for i3 in range(nn.Elements.Count):
            e = nn.Elements.Item(i3)
            try:
                v = float(e.Value) if e.Value is not None else 0.0
                if abs(v) > 0.001:
                    d[e.Name] = v
            except Exception:
                pass
    except Exception:
        pass
    return d


L = ['=== 关键面板消息 ===']
for i2, x in enumerate(MSGS):
    if re.search(r'SEVERE|Terminal|Severe|Errors|Warnings|completed|FLASH', x):
        L.append('%3d | %s' % (i2, x[:165]))

IN = ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-132', 'S-210']
OUTL = ['S-130', 'S-131', 'S-211', 'S-212', 'S-213', 'S-115', 'S-119',
        'S-122', 'S-216', 'S-217', 'S-218', 'S-405C', 'S-406', 'S-209']
vi = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in IN)
vo = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in OUTL)
L.append('')
L.append('=== 总物料平衡 ===  进 %.2f  出 %.2f  偏差 %.4f %%' % (vi, vo, (vo - vi) / vi * 100))

p = comp('S-209')
tt = sum(p.values())
L.append('')
L.append('=== 产品 S-209 = %.2f kg/h ===' % tt)
for k, v in sorted(p.items(), key=lambda x: -x[1]):
    L.append('    %-8s %9.2f  %6.2f wt%%' % (k, v, v / tt * 100 if tt else 0))
L.append('  折年产 = %.0f t/a' % (tt / 1000 * 7200))
nm = p.get('NAM', 0)
mp3 = comp('S-101').get('3-MP', 0)
L.append('  3-MP 单耗 = %.1f kg/t    总收率 = %.2f%%'
         % (mp3 / nm * 1000, nm / 122.13 / (mp3 / 93.13) * 100))

L.append('')
L.append('=== 全流股组成（论文用）===')
for s in ['S-201', 'S-203', 'S-211', 'S-212', 'S-213', 'S-401', 'S-402', 'S-214',
          'S-205', 'S-403', 'S-206', 'S-404', 'S-405', 'S-405B', 'S-405C', 'S-208',
          'S-406', 'S-209', 'S-121', 'S-117', 'S-117M', 'S-117T', 'S-119', 'S-122']:
    c = comp(s)
    tot2 = sum(c.values())
    L.append('  %-8s %10s kg/h | %s' % (s, g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s),
                                        ', '.join('%s %.2f' % (k, v) for k, v in
                                                  sorted(c.items(), key=lambda x: -x[1])[:5])))

L.append('')
L.append('=== 后处理收率核算 ===')
nam201 = comp('S-201').get('NAM', 0)
nam209 = p.get('NAM', 0)
L.append('  进精制工段烟酰胺（S-201） = %.2f kg/h' % nam201)
L.append('  产品烟酰胺（S-209）       = %.2f kg/h' % nam209)
if nam201 > 0:
    L.append('  精制段总收率 = %.4f' % (nam209 / nam201))
L.append('  母液 S-405 = %.2f kg/h；循环 S-405B = %.2f；排放 S-405C = %.2f'
         % (float(g(r'\Data\Streams\S-405\Output\MASSFLMX\MIXED') or 0),
            float(g(r'\Data\Streams\S-405B\Output\MASSFLMX\MIXED') or 0),
            float(g(r'\Data\Streams\S-405C\Output\MASSFLMX\MIXED') or 0)))

L.append('')
L.append('=== 三效蒸发 ===')
for s in ['S-203', 'S-401', 'S-402', 'S-214', 'S-205', 'S-403', 'S-206']:
    c = comp(s)
    tt2 = sum(c.values())
    L.append('  %-7s %9.2f kg/h  NAM %6.2f wt%%  水 %8.2f' % (
        s, tt2, c.get('NAM', 0) / tt2 * 100 if tt2 else 0, c.get('H2O', 0)))

try:
    doc.SaveAs(OUT.replace('.bkp', '_s.bkp'))
    L.append('')
    L.append('已保存')
except Exception as ex:
    L.append('保存失败: %s' % ex)
try:
    doc.Close()
except Exception:
    pass

open(r'D:\<化工工作区>\_probe\s13.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('DONE')
