# -*- coding: utf-8 -*-
"""S8：水解转化率对齐论文 99%、干燥器提温降水分，出最终结果与全流程校核"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_s7.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_s8.bkp'
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


# 1) R-501 水解转化率 .985 -> .99（论文口径：水解收率 99%）
def fn501(seg):
    seg2, n = re.subn(r'(KEY-CID\s*=\s*"3-CP"\s*CONV\s*=\s*)[\d.]+', r'\g<1>.99', seg)
    print('   R-501 CONV 替换 %d 处' % n)
    return seg2


t = repl_block(t, 'R-501', fn501)
i = t.find('BLOCK RSTOIC "R-501"')
seg501 = t[i:i + 2000]
print('1) R-501 段 CONV 片段:', re.findall(r'CONV = [\d.]+', seg501))


# 2) 干燥器提温到 130 ℃（水分降到 0.5% 以下）
def fn601(seg):
    return re.sub(r'PARAM TEMP = [\d.]+ <22> <4>', 'PARAM TEMP = 130. <22> <4>', seg, count=1)


t = repl_block(t, 'D-601', fn601)
print('2) D-601 -> 130 ℃ ✓')

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
while time.time() - t0 < 540:
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
    if re.search(r'ERROR|SEVERE|WARNING|Terminal|Severe|Errors|Warnings|completed|BYPASSED|NOT CONVERG', x):
        L.append('%3d | %s' % (i2, x[:170]))

IN = ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-132', 'S-210']
OUTL = ['S-130', 'S-131', 'S-211', 'S-212', 'S-213', 'S-115', 'S-117', 'S-119',
        'S-122', 'S-216', 'S-217', 'S-218', 'S-405C', 'S-406', 'S-209']
vi = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in IN)
vo = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in OUTL)
L.append('')
L.append('=== 总物料平衡 ===')
L.append('  进料 = %.2f kg/h   出料 = %.2f kg/h   偏差 = %.3f %%' % (vi, vo, (vo - vi) / vi * 100))

c209 = comp('S-209')
t209 = sum(c209.values())
L.append('')
L.append('=== 产品 S-209（干燥烟酰胺）===')
L.append('  总质量 %.2f kg/h' % t209)
for k, v in sorted(c209.items(), key=lambda x: -x[1]):
    L.append('    %-8s %9.2f kg/h  %6.2f wt%%' % (k, v, v / t209 * 100 if t209 else 0))
L.append('  折年产（7200 h）= %.0f t/a' % (t209 / 1000 * 7200))

c121 = comp('S-121')
t121 = sum(c121.values())
L.append('')
L.append('=== 进水解釜的 3-CP（S-121）===')
L.append('  总质量 %.2f kg/h' % t121)
for k, v in sorted(c121.items(), key=lambda x: -x[1]):
    L.append('    %-8s %9.2f kg/h  %6.2f wt%%' % (k, v, v / t121 * 100 if t121 else 0))

c101 = comp('S-101')
t101 = sum(c101.values())
L.append('')
L.append('=== 原料 S-101 ===')
L.append('  总质量 %.2f kg/h' % t101)
for k, v in sorted(c101.items(), key=lambda x: -x[1]):
    L.append('    %-8s %9.2f kg/h' % (k, v))
nm = c209.get('NAM', 0)
mp3 = c101.get('3-MP', 0)
if nm > 0:
    L.append('  3-甲基吡啶单耗 = %.1f kg/t 烟酰胺' % (mp3 / nm * 1000))
    L.append('  总收率（3-MP→烟酰胺，摩尔）= %.2f%%'
             % (nm / 122.13 / (mp3 / 93.13) * 100))

L.append('')
L.append('=== 三效蒸发 ===')
for s in ['S-203', 'S-401', 'S-402', 'S-214', 'S-205', 'S-403', 'S-206']:
    c = comp(s)
    tt = sum(c.values())
    L.append('  %-7s %9.2f kg/h   NAM %7.2f wt%%   水 %8.2f' % (
        s, tt, c.get('NAM', 0) / tt * 100 if tt else 0, c.get('H2O', 0)))

L.append('')
L.append('=== 主要流股汇总 ===')
for s in ['S-104', 'S-106', 'S-110', 'S-113', 'S-114', 'S-115', 'S-117', 'S-118',
          'S-119', 'S-120', 'S-121', 'S-122', 'S-201', 'S-203', 'S-211', 'S-212',
          'S-213', 'S-401', 'S-403', 'S-404', 'S-405', 'S-405B', 'S-405C', 'S-208',
          'S-209', 'S-130', 'S-131']:
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

open(r'D:\<化工工作区>\_probe\s8.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('DONE')
