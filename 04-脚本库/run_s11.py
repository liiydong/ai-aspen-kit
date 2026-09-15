# -*- coding: utf-8 -*-
"""S11：T-402 采出比 0.040->0.070；M-504 改为按 S-117M 定分数（甲苯不进反应器）"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_s10.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_s11.bkp'
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


# ---------- 1) T-402 采出比 ----------
s0, s1 = find_block_span(t, 'T-402')
seg = t[s0:s1]
print('--- T-402 COL-SPECS 片段 ---')
for m in re.finditer(r'"COL-SPECS"[^\\]{0,160}', re.sub(r'\s+', ' ', seg)):
    print('   ', m.group()[:160])
print('   所有 D 相关字段:', re.findall(r'[A-Z-]*D[F:.-]*\s*=\s*[\d.]+', re.sub(r'\s+', ' ', seg))[:8])


def fn402(s):
    pats = [r'(BASIS-DF\s*=\s*)[\d.]+', r'(BASIS-D:F\s*=\s*)[\d.]+', r'(D:F\s*=\s*)[\d.]+',
            r'(BASIS-MASSDF\s*=\s*)[\d.]+', r'(D-F\s*=\s*)[\d.]+']
    for p in pats:
        s2, n = re.subn(p, lambda m: m.group(1) + '0.070', s)
        if n:
            print('   命中模式 %s（%d 处）' % (p, n))
            return s2
    raise RuntimeError('未找到 T-402 的 D:F 字段')


t = repl_block(t, 'T-402', fn402)
print('1) T-402 采出比 -> 0.070 ✓')

# ---------- 2) M-504 改为按 S-117M 定分数 ----------
ci = t.find('? COMPONENTS MAIN ?')
cj = t.find('? COMPONENTS "COMP-LIST"', ci)
cids = [c for c in re.findall(r'CID = "?([A-Za-z0-9-]+)"?', t[ci:cj]) if c != 'SUBSTREAM']
fr = {'3-MP': 0.98, '4-MP': 0.98}   # 其余全部（含 TOL）留给 S-117T
recs = ['PARAM-STREAM = "S-117M" SUBSTREAM = MIXED COMPS = "%s" FRACS = %s <0> <0> '
        % (c, fr.get(c, 0.0)) for c in cids]
seg = ('? BLOCK SEP "M-504" ? ; "METCBAR_MOLE" ; ; ICON1 ; \\ PARAM1 PRES1 = 1.0 <20> <5> \\ \\ '
       'PARAM ') + ' /  '.join(recs) + ' \\ '
t = repl_block(t, 'M-504', lambda s: seg)
print('2) M-504 改为按 S-117M 定分数（甲苯 0 -> 全部走 S-117T 回萃取塔）✓')

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
    if re.search(r'SEVERE|Terminal|Severe|Errors|Warnings|completed|BYPASSED|FLASH', x):
        L.append('%3d | %s' % (i2, x[:165]))

IN = ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-132', 'S-210']
OUTL = ['S-130', 'S-131', 'S-211', 'S-212', 'S-213', 'S-115', 'S-117T', 'S-119',
        'S-122', 'S-216', 'S-217', 'S-218', 'S-405C', 'S-406', 'S-209']
vi = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in IN)
vo = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in OUTL)
L.append('')
L.append('=== 总物料平衡 ===  进 %.2f  出 %.2f  偏差 %.3f %%' % (vi, vo, (vo - vi) / vi * 100))

p = comp('S-209')
tt = sum(p.values())
L.append('')
L.append('=== 产品 S-209 = %.2f kg/h ===' % tt)
for k, v in sorted(p.items(), key=lambda x: -x[1]):
    L.append('    %-8s %9.2f  %6.2f wt%%' % (k, v, v / tt * 100 if tt else 0))
L.append('  折年产 = %.0f t/a' % (tt / 1000 * 7200))
nm = p.get('NAM', 0)
mp3 = comp('S-101').get('3-MP', 0)
if nm > 0:
    L.append('  3-MP 单耗 = %.1f kg/t    总收率 = %.2f%%'
             % (mp3 / nm * 1000, nm / 122.13 / (mp3 / 93.13) * 100))

L.append('')
L.append('=== 回收点与产品塔 ===')
for s in ['S-117', 'S-117M', 'S-117T', 'S-116', 'S-118', 'S-119', 'S-120', 'S-121', 'S-122']:
    c = comp(s)
    tot2 = sum(c.values())
    L.append('  --- %s (%s kg/h) ---' % (s, g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s)))
    for k, v in sorted(c.items(), key=lambda x: -x[1]):
        if tot2 and v / tot2 > 0.002:
            L.append('      %-8s %9.2f  %6.2f wt%%' % (k, v, v / tot2 * 100))
    if not c:
        L.append('      (无数据)')

L.append('')
L.append('=== 主流程关键流股 ===')
for s in ['S-104', 'S-106', 'S-110', 'S-113', 'S-114', 'S-115', 'S-121', 'S-209',
          'S-201', 'S-203', 'S-405', 'S-405C', 'S-130']:
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

open(r'D:\<化工工作区>\_probe\s11.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('DONE')
