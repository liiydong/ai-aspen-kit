# -*- coding: utf-8 -*-
"""S7：补全设备（冷却器/三效冷凝器/母液循环+排放/废水泵）+ 修正干燥器 + 标定浓缩"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_s6.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_s7.bkp'
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


def fs_entry_span(x, bid):
    ms = list(re.finditer(r'BLOCK\s+BLKID\s*=\s*"([A-Za-z0-9_.-]+)"', x))
    for k, m in enumerate(ms):
        if m.group(1) == bid:
            s0 = m.start()
            s1 = ms[k + 1].start() if k + 1 < len(ms) else x.find('"DEF-STREAM"', s0)
            return s0, s1
    raise RuntimeError('FLOWSHEET 未找到 ' + bid)


def fs_set_in(x, bid, new_in):
    i, j = fs_entry_span(x, bid)
    seg2, n = re.subn(r'IN\s*=\s*\([^)]*\)', 'IN = ( %s )' % new_in, x[i:j], count=1)
    assert n == 1, bid
    return x[:i] + seg2 + x[j:]


def set_frac(x, bid, cid, val):
    def fn(seg):
        pat = re.compile(r'(COMPS\s*=\s*"%s"\s*FRACS\s*=\s*)[\d.]+' % re.escape(cid))
        return pat.sub(lambda m: m.group(1) + str(val), seg)
    return repl_block(x, bid, fn)


def set_heater(x, bid, temp, pres):
    def fn(seg):
        seg2 = re.sub(r'PARAM\s+TEMP\s*=\s*[\d.]+\s*<22>\s*<\d+>\s*PRES\s*=\s*[\d.]+\s*<20>\s*<\d+>',
                      'PARAM TEMP = %s <22> <4> PRES = %s <20> <5>' % (temp, pres), seg, count=1)
        if seg2 == seg:
            raise RuntimeError('set_heater 未匹配 ' + bid)
        return seg2
    return repl_block(x, bid, fn)


# ============ 1) 注册表新增 7 个块 ============
NEW = [('E-501', 'Heater', 'HEATER'), ('E-604', 'Heater', 'HEATER'),
       ('E-605', 'Heater', 'HEATER'), ('E-606', 'Heater', 'HEATER'),
       ('P-301' and 'P-501', 'Pump', 'PUMP'), ('M-503', 'Sep', 'SEP'),
       ('P-302', 'Pump', 'PUMP')]
m = re.search(r';\s*\n(\d+)\s*\n>VERSION 0', t)
n_old = int(m.group(1))
A_STR = t.find('? SETUP MAIN ?')
t = t[:A_STR] + ''.join('>VERSION 0\n%s\n%s\nBuilt-In\n%s\n' % e for e in NEW) + t[A_STR:]
t = t[:m.start()] + '; \n%d\n>VERSION 0' % (n_old + len(NEW)) + t[m.end():]
print('1) 注册表 %d -> %d ✓' % (n_old, n_old + len(NEW)))

# ============ 2) FLOWSHEET ============
A_FS = t.find('"DEF-STREAM"')


def ent(bid, bt, mt, ins_, outs_):
    return ('BLOCK BLKID = "%s" BLKTYPE = "%s" MDLTYPE = "%s" IN = ( %s ) OUT = ( %s ) \\ \\ '
            % (bid, bt, mt, ins_, outs_))


ins = (ent('E-501', 'HEATER', 'Heater', '"S-201" M0-1', '"S-215" M0-1')
       + ent('E-604', 'HEATER', 'Heater', '"S-402" M0-1', '"S-216" M0-1')
       + ent('E-605', 'HEATER', 'Heater', '"S-205" M0-1', '"S-217" M0-1')
       + ent('E-606', 'HEATER', 'Heater', '"S-206" M0-1', '"S-218" M0-1')
       + ent('P-501', 'PUMP', 'Pump', '"S-405" M0-1', '"S-405A" M0-1')
       + ent('M-503', 'SEP', 'Sep', '"S-405A" M0-1', '"S-405B" M0-1 "S-405C" M0-1')
       + ent('P-302', 'PUMP', 'Pump', '"S-213" M0-1', '"S-213A" M0-1'))
t = t[:A_FS] + wr(ins) + t[A_FS:]
print('2a) FLOWSHEET 追加 7 条 ✓')

# M-501 进料改为 S-215 + S-210 + 母液循环 S-405B
t = fs_set_in(t, 'M-501', '"S-215" M0-1 "S-210" M0-1 "S-405B" M0-1')
print('2b) M-501 接入母液循环 S-405B ✓')

# ============ 3) 新块段落 ============
ci = t.find('? COMPONENTS MAIN ?')
cj = t.find('? COMPONENTS "COMP-LIST"', ci)
cids = re.findall(r'CID = "?([A-Za-z0-9-]+)"?', t[ci:cj])
cids = [c for c in cids if c not in ('SUBSTREAM',)]


def heater_para(bid, desc, temp, pres):
    return ('? BLOCK HEATER "%s" ? ; "METCBAR_MOLE" ; ; HEATER ; \\ DESCRIPTION '
            'DESCRIPTION = "%s" \\ \\ PARAM TEMP = %s <22> <4> PRES = %s <20> <5> \\ '
            % (bid, desc, temp, pres))


def pump_para(bid, desc, pres):
    return ('? BLOCK PUMP "%s" ? ; "METCBAR_MOLE" ; ; ICON1 ; \\ DESCRIPTION '
            'DESCRIPTION = "%s" \\ \\ PARAM PRES = %s <20> <5> EFF = 0.7 <0> <0> OPT-SPEC = PRES \\ '
            % (bid, desc, pres))


def sep_para(bid, out_s, fr):
    recs = ['PARAM-STREAM = "%s" SUBSTREAM = MIXED COMPS = "%s" FRACS = %s <0> <0> ' % (out_s, c, fr.get(c, 0.0))
            for c in cids]
    return ('? BLOCK SEP "%s" ? ; "METCBAR_MOLE" ; ; ICON1 ; \\ PARAM1 PRES1 = 1.0 <20> <5> \\ \\ '
            'PARAM ' % bid) + ' /  '.join(recs) + ' \\ '


S = [wr(heater_para('E-501', 'Hydrolysate cooler', '80.', '1.0')),
     wr(heater_para('E-604', 'First effect condenser', '45.', '1.0')),
     wr(heater_para('E-605', 'Second effect condenser', '45.', '0.47')),
     wr(heater_para('E-606', 'Third effect condenser', '45.', '0.20')),
     wr(pump_para('P-501', 'Mother liquor pump', '1.5')),
     wr(pump_para('P-302', 'Wastewater pump', '1.5')),
     wr(sep_para('M-503', 'S-405C', {c: 0.12 for c in cids}))]
A_BLOCK = t.find('? BLOCK')
t = t[:A_BLOCK] + '\n'.join(S) + t[A_BLOCK:]
print('3) 7 个新块段落 ✓（组分 %d 个）' % len(cids))

# ============ 4) 离心机分离分数（单程结晶 ~75% 析出 → 25% 进母液）============
for cid, v in [('NAM', 0.25), ('H2O', 0.90), ('NAC', 0.5), ('3-CP', 0.95),
               ('3-MP', 0.95), ('NH3', 0.0), ('O2', 0.0), ('N2', 0.0), ('4-CP', 0.0),
               ('TOL', 0.0), ('CO2', 0.0), ('HCN', 0.0), ('CO', 0.0), ('AIR', 0.0)]:
    t = set_frac(t, 'M-601', cid, v)
print('4) M-601 单程结晶收率 75%（NAM 25% 进母液）✓')

# ============ 5) 干燥器：降水分 ============
t = set_heater(t, 'D-601', '100.', '0.05')
print('5) D-601 干燥器 -> 100 ℃ / 0.05 bar ✓')

# ============ 6) 三效蒸发汽化率（蒸汽再压缩回收后浓缩到约 60 wt%）============
for bid, temp, pres, vf in [('E-601', '100.', '1.0', '0.30'),
                            ('E-602', '80.', '0.47', '0.35'),
                            ('E-603', '60.', '0.20', '0.40')]:
    def fn(seg, temp=temp, pres=pres, vf=vf):
        seg = re.sub(r'PARAM\s+TEMP\s*=\s*[\d.]+\s*<22>\s*<\d+>\s*PRES\s*=\s*[\d.]+\s*<20>\s*<\d+>'
                     r'(\s*VFRAC\s*=\s*[\d.]+\s*<0>\s*<0>)?(\s*SPEC-OPT\s*=\s*\w+)?',
                     'PARAM TEMP = %s <22> <4> PRES = %s <20> <5> VFRAC = %s <0> <0> SPEC-OPT = PV'
                     % (temp, pres, vf), seg, count=1)
        return seg
    t = repl_block(t, bid, fn)
print('6) 三效蒸发 汽化率 0.30 / 0.35 / 0.40 ✓')

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


L = []
L.append('=== 关键面板消息 (%d 条) ===' % len(MSGS))
for i2, x in enumerate(MSGS):
    if re.search(r'ERROR|SEVERE|WARNING|Terminal|Severe|Errors|Warnings|completed|BYPASSED|NOT CONVERG|Tear|tear', x):
        L.append('%3d | %s' % (i2, x[:170]))

L.append('')
ALL = ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-132', 'S-210',
       'S-104', 'S-106', 'S-108', 'S-109', 'S-110', 'S-112', 'S-113', 'S-114',
       'S-114A', 'S-115', 'S-116', 'S-116A', 'S-117', 'S-118', 'S-119', 'S-120',
       'S-121', 'S-122', 'S-130', 'S-131',
       'S-201', 'S-215', 'S-202', 'S-203', 'S-211', 'S-212', 'S-213', 'S-213A',
       'S-401', 'S-402', 'S-216', 'S-214', 'S-205', 'S-217', 'S-403', 'S-206', 'S-218',
       'S-404', 'S-405', 'S-405A', 'S-405B', 'S-405C', 'S-208', 'S-406', 'S-209']
L.append('=== 流股质量流量 kg/h ===')
for s in ALL:
    L.append('  %-8s %-12s' % (s, g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s)))

IN = ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-132', 'S-210']
OUTL = ['S-130', 'S-131', 'S-211', 'S-212', 'S-213', 'S-115', 'S-117', 'S-119',
        'S-122', 'S-216', 'S-217', 'S-218', 'S-405C', 'S-406', 'S-209']


def tot(names):
    v = 0.0
    miss = []
    for s in names:
        x = g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s)
        if x is None:
            miss.append(s)
        else:
            v += float(x)
    return v, miss


vi, mi = tot(IN)
vo, mo = tot(OUTL)
L.append('')
L.append('=== 总物料平衡 ===')
L.append('  进料 = %.2f   出料 = %.2f   偏差 = %.2f kg/h (%.3f%%)  缺:%s/%s'
         % (vi, vo, vo - vi, (vo - vi) / vi * 100 if vi else 0, mi, mo))

L.append('')
L.append('=== 关键流股组成 ===')
for s in ['S-203', 'S-401', 'S-214', 'S-403', 'S-404', 'S-405', 'S-405B', 'S-405C',
          'S-208', 'S-209', 'S-212', 'S-213']:
    nn = doc.Tree.FindNode(r'\Data\Streams\%s\Output\MASSFLOW3' % s)
    if nn is None:
        L.append('  %s: 无数据' % s)
        continue
    L.append('  --- %s (总 %s) ---' % (s, g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s)))
    try:
        for i3 in range(nn.Elements.Count):
            e = nn.Elements.Item(i3)
            try:
                if e.Value and abs(float(e.Value)) > 0.05:
                    L.append('      %-8s %9.2f' % (e.Name, float(e.Value)))
            except Exception:
                pass
    except Exception as ex:
        L.append('      读取失败 %s' % ex)

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

open(r'D:\<化工工作区>\_probe\s7.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('DONE')
