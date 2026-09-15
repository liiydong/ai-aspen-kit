# -*- coding: utf-8 -*-
"""S10：加 M-504 甲苯回收器（甲苯->萃取塔 / 3-MP->反应器），修正 S-117 回收路线"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_s9.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_s10.bkp'
t = open(SRC, encoding='utf-8', errors='ignore').read()
print('源大小', len(t))


def wr(s, w=74):
    out, line = [], ''
    for tok in s.split(' '):
        if len(line) + len(tok) + 1 > w:
            out.append(line)
            line = tok
        else:
            line = (line + ' ' + tok) if line else tok
    if line:
        out.append(line)
    return '\n'.join(out)


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


# 1) 注册表新增 M-504
NEW = [('M-504', 'Sep', 'SEP')]
m = re.search(r';\s*\n(\d+)\s*\n>VERSION 0', t)
n_old = int(m.group(1))
A_STR = t.find('? SETUP MAIN ?')
t = t[:A_STR] + ''.join('>VERSION 0\n%s\n%s\nBuilt-In\n%s\n' % e for e in NEW) + t[A_STR:]
t = t[:m.start()] + '; \n%d\n>VERSION 0' % (n_old + 1) + t[m.end():]
print('1) 注册表 %d -> %d ✓' % (n_old, n_old + 1))

# 2) FLOWSHEET：新增 M-504；M-101 吃 S-117M；T-301 吃 S-117T
A_FS = t.find('"DEF-STREAM"')
ent = ('BLOCK BLKID = "M-504" BLKTYPE = "SEP" MDLTYPE = "Sep" IN = ( "S-117" M0-1 ) '
       'OUT = ( "S-117M" M0-1 "S-117T" M0-1 ) \\ \\ ')
t = t[:A_FS] + wr(ent) + t[A_FS:]
t = fs_set_in(t, 'M-101', '"S-101" M0-1 "S-126" M0-1 "S-128" M0-1 "S-117M" M0-1')
t = fs_set_in(t, 'T-301', '"S-112" M0-1 "S-111" M0-1 "S-117T" M0-1')
print('2) FLOWSHEET：M-504 接入；M-101<-S-117M；T-301<-S-117T ✓')

# 3) M-504 段（Sep）：甲苯/3-CP/4-CP -> S-117T；3-MP -> S-117M
ci = t.find('? COMPONENTS MAIN ?')
cj = t.find('? COMPONENTS "COMP-LIST"', ci)
cids = [c for c in re.findall(r'CID = "?([A-Za-z0-9-]+)"?', t[ci:cj]) if c != 'SUBSTREAM']
fr = {'TOL': 0.98, '3-CP': 0.90, '4-CP': 0.90, 'H2O': 0.50, '3-MP': 0.02}
recs = ['PARAM-STREAM = "S-117T" SUBSTREAM = MIXED COMPS = "%s" FRACS = %s <0> <0> '
        % (c, fr.get(c, 0.01)) for c in cids]
seg = ('? BLOCK SEP "M-504" ? ; "METCBAR_MOLE" ; ; ICON1 ; \\ PARAM1 PRES1 = 1.0 <20> <5> \\ \\ '
       'PARAM ') + ' /  '.join(recs) + ' \\ '
A_BLOCK = t.find('? BLOCK')
t = t[:A_BLOCK] + wr(seg) + '\n' + t[A_BLOCK:]
print('3) M-504 段 ✓（%d 组分）' % len(cids))

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
    if re.search(r'ERROR|SEVERE|FLASH|Terminal|Severe|Errors|Warnings|completed|BYPASSED|NOT CONVERG', x):
        L.append('%3d | %s' % (i2, x[:165]))

IN = ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-132', 'S-210']
OUTL = ['S-130', 'S-131', 'S-211', 'S-212', 'S-213', 'S-115', 'S-117', 'S-119',
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
L.append('=== 回收与产品塔流股 ===')
for s in ['S-117', 'S-117M', 'S-117T', 'S-122', 'S-118', 'S-119', 'S-121', 'S-106']:
    c = comp(s)
    tot2 = sum(c.values())
    L.append('  --- %s (%s kg/h) ---' % (s, g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s)))
    for k, v in sorted(c.items(), key=lambda x: -x[1]):
        if tot2 and v / tot2 > 0.002:
            L.append('      %-8s %9.2f  %6.2f wt%%' % (k, v, v / tot2 * 100))
    if not c:
        L.append('      (无数据)')

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

open(r'D:\<化工工作区>\_probe\s10.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('DONE')
