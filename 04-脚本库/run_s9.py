# -*- coding: utf-8 -*-
"""S9：接入 S-117（未反应 3-MP）回反应器的回收循环，补齐产量"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_s8_s.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_s9.bkp'
t = open(SRC, encoding='utf-8', errors='ignore').read()
print('源大小', len(t))


def fs_entry_span(x, bid):
    ms = list(re.finditer(r'BLOCK\s+BLKID\s*=\s*"([A-Za-z0-9_.-]+)"', x))
    for k, m in enumerate(ms):
        if m.group(1) == bid:
            s0 = m.start()
            s1 = ms[k + 1].start() if k + 1 < len(ms) else x.find('"DEF-STREAM"', s0)
            return s0, s1
    raise RuntimeError(bid)


def fs_set_in(x, bid, new_in):
    i, j = fs_entry_span(x, bid)
    seg2, n = re.subn(r'IN\s*=\s*\([^)]*\)', 'IN = ( %s )' % new_in, x[i:j], count=1)
    assert n == 1, bid
    return x[:i] + seg2 + x[j:]


def section_span(x, sid):
    m = re.search(r'\?\s*STREAM\s+MATERIAL\s+"%s"' % re.escape(sid), x)
    assert m, '未找到流股段 ' + sid
    i = m.start()
    nxt = re.search(r'\?\s*(?:STREAM\s+MATERIAL|BLOCK)\s', x[i + 10:])
    j = i + 10 + nxt.start() if nxt else len(x)
    return i, j


def scale_stream(x, sid, f):
    i, j = section_span(x, sid)
    seg = x[i:j]
    seg2, n1 = re.subn(r'((?:FLOW|TOTAL) = )([\d.eE+-]+)',
                       lambda m: '%s%.6f' % (m.group(1), float(m.group(2)) * f), seg)
    assert n1 > 0, sid
    print('   %s 缩放 ×%.3f（%d 个数值）' % (sid, f, n1))
    return x[:i] + seg2 + x[j:]


# 1) S-117 -> M-101 反应器进料（回收未反应 3-甲基吡啶）
t = fs_set_in(t, 'M-101', '"S-101" M0-1 "S-126" M0-1 "S-128" M0-1 "S-117" M0-1')
print('1) S-117 -> M-101 回收循环接入 ✓')

# 2) 反应器与水解工段同步放大约 3%（3-MP 总进料提高 2.9%）
F = 1.03
t = scale_stream(t, 'S-102', F)
t = scale_stream(t, 'S-107', F)
t = scale_stream(t, 'S-111', F)
t = scale_stream(t, 'S-124', F)
print('2) S-102 / S-107 / S-111 / S-124 同步放大约 3% ✓')

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


L = []
L.append('=== 关键面板消息 ===')
for i2, x in enumerate(MSGS):
    if re.search(r'ERROR|SEVERE|WARNING|Terminal|Severe|Errors|Warnings|completed|BYPASSED|NOT CONVERG|tear', x):
        L.append('%3d | %s' % (i2, x[:170]))

IN = ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-132', 'S-210']
OUTL = ['S-130', 'S-131', 'S-211', 'S-212', 'S-213', 'S-115', 'S-117', 'S-119',
        'S-122', 'S-216', 'S-217', 'S-218', 'S-405C', 'S-406', 'S-209']
vi = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in IN)
vo = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in OUTL)
L.append('')
L.append('=== 总物料平衡 ===')
L.append('  进料 = %.2f   出料 = %.2f   偏差 = %.3f %%' % (vi, vo, (vo - vi) / vi * 100))

p209 = comp('S-209')
tot9 = sum(p209.values())
L.append('')
L.append('=== 产品 S-209 ===')
L.append('  总质量 %.2f kg/h' % tot9)
for k, v in sorted(p209.items(), key=lambda x: -x[1]):
    L.append('    %-8s %9.2f kg/h  %6.2f wt%%' % (k, v, v / tot9 * 100 if tot9 else 0))
L.append('  折年产（7200 h）= %.0f t/a' % (tot9 / 1000 * 7200))

c101 = comp('S-101')
mp3 = c101.get('3-MP', 0)
nm = p209.get('NAM', 0)
L.append('')
L.append('=== 原料与单耗 ===')
L.append('  S-101 3-MP = %.2f kg/h；S-102 NH3 = %s kg/h；S-124 水解水 = %s'
         % (mp3, g(r'\Data\Streams\S-102\Output\MASSFLMX\MIXED'), g(r'\Data\Streams\S-124\Output\MASSFLMX\MIXED')))
if nm > 0:
    L.append('  3-MP 单耗 = %.1f kg/t 烟酰胺' % (mp3 / nm * 1000))
    L.append('  总收率（3-MP->烟酰胺，摩尔）= %.2f%%' % (nm / 122.13 / (mp3 / 93.13) * 100))

L.append('')
L.append('=== 回收点流股 ===')
for s in ['S-117', 'S-122', 'S-118', 'S-119', 'S-120', 'S-121', 'S-116', 'S-114']:
    c = comp(s)
    tt = sum(c.values())
    L.append('  --- %s (总 %s kg/h) ---' % (s, g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s)))
    for k, v in sorted(c.items(), key=lambda x: -x[1]):
        if v / tt * 100 > 0.3 if tt else False:
            L.append('      %-8s %9.2f  %6.2f wt%%' % (k, v, v / tt * 100))

L.append('')
L.append('=== 关键流股流量 ===')
for s in ['S-104', 'S-106', 'S-109', 'S-110', 'S-113', 'S-114', 'S-115', 'S-116',
          'S-117', 'S-118', 'S-119', 'S-120', 'S-121', 'S-122', 'S-130', 'S-131',
          'S-201', 'S-203', 'S-209', 'S-211', 'S-405', 'S-405C']:
    L.append('  %-8s %10s' % (s, g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s)))

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

open(r'D:\<化工工作区>\_probe\s9.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('DONE')
