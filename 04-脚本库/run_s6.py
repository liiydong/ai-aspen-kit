# -*- coding: utf-8 -*-
"""S6：修正三效蒸发链（液相走向 + 汽化率规格 PV）、干燥器端口、脱色/离心分离分数"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_s5.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_s6.bkp'
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


def set_flash(x, bid, temp, pres):
    """把 Flash2 的 PARAM 改成 温度+压力 规格"""
    def fn(seg):
        new = ('PARAM TEMP = %s <22> <4> PRES = %s <20> <5>' % (temp, pres))
        seg2 = re.sub(r'PARAM\s+TEMP\s*=\s*[\d.]+\s*<22>\s*<\d+>\s*PRES\s*=\s*[\d.]+\s*<20>\s*<\d+>',
                      new, seg, count=1)
        if seg2 == seg:
            raise RuntimeError('set_flash 未匹配 ' + bid)
        return seg2
    return repl_block(x, bid, fn)


def set_vfrac(x, bid, temp, pres, vfrac):
    """Flash2 改成 压力+汽化率 规格（SPEC-OPT = PV）"""
    def fn(seg):
        new = ('PARAM TEMP = %s <22> <4> PRES = %s <20> <5> VFRAC = %s <0> <0> SPEC-OPT = PV'
               % (temp, pres, vfrac))
        seg2 = re.sub(r'PARAM\s+TEMP\s*=\s*[\d.]+\s*<22>\s*<\d+>\s*PRES\s*=\s*[\d.]+\s*<20>\s*<\d+>',
                      new, seg, count=1)
        if seg2 == seg:
            raise RuntimeError('set_vfrac 未匹配 ' + bid)
        return seg2
    return repl_block(x, bid, fn)


def set_frac(x, bid, cid, val):
    """改 Sep 块里某个组分流向第一股产品的分数"""
    def fn(seg):
        pat = re.compile(r'(COMPS\s*=\s*"%s"\s*FRACS\s*=\s*)[\d.]+' % re.escape(cid))
        seg2, n = pat.subn(lambda m: m.group(1) + str(val), seg)
        return seg2
    return repl_block(x, bid, fn)


def fs_entry_span(x, bid):
    ms = list(re.finditer(r'BLOCK\s+BLKID\s*=\s*"([A-Za-z0-9_.-]+)"', x))
    for k, m in enumerate(ms):
        if m.group(1) == bid:
            s0 = m.start()
            if k + 1 < len(ms):
                s1 = ms[k + 1].start()
            else:
                s1 = x.find('"DEF-STREAM"', s0)
            return s0, s1
    raise RuntimeError('FLOWSHEET 未找到 ' + bid)


def fs_set_in(x, bid, new_in):
    i, j = fs_entry_span(x, bid)
    seg2, n = re.subn(r'IN\s*=\s*\([^)]*\)', 'IN = ( %s )' % new_in, x[i:j], count=1)
    assert n == 1, bid
    return x[:i] + seg2 + x[j:]


def fs_set_out(x, bid, new_out):
    i, j = fs_entry_span(x, bid)
    seg2, n = re.subn(r'OUT\s*=\s*\([^)]*\)', 'OUT = ( %s )' % new_out, x[i:j], count=1)
    assert n == 1, bid
    return x[:i] + seg2 + x[j:]


# ---------- 1) 三效蒸发：液相走向修正 ----------
# 端口语义（COM 实测）：Flash2  M0-1 = V(气相)、M1-2 = L(液相)
#   E-601: V=S-402  L=S-401  -> E-602 应吃 S-401 ✓（已是）
#   E-602: V=S-205  L=S-214  -> E-603 应吃 S-214（原接 S-205 气相）✗
#   E-603: V=S-206  L=S-403  -> C-601 应吃 S-403（原接 S-206 气相）✗
#   C-601: V=S-207  L=S-404  -> M-601 应吃 S-404（原接 S-207 气相）✗
t = fs_set_in(t, 'E-603', '"S-214" M0-1')
t = fs_set_in(t, 'C-601', '"S-403" M0-1')
t = fs_set_in(t, 'M-601', '"S-404" M0-1')
print('1) 三效蒸发链液相走向修正 ✓  E-603<-S-214, C-601<-S-403, M-601<-S-404')

# ---------- 2) 干燥器端口：气相 -> M0-1，干品 -> M1-2 ----------
t = fs_set_out(t, 'D-601', '"S-406" M0-1 "S-209" M1-2')
print('2) D-601 端口修正 ✓  S-406(水汽)=M0-1, S-209(干品)=M1-2')

# ---------- 3) 三效蒸发汽化率（按浓缩到 70% NAM 计算） ----------
# 进料 4151.7 kg/h（NAM 1379.7 + 水 2754.1）；目标出料 NAM 70 wt%
# 需蒸发水 2162.8 kg/h，三效均分 720.9 kg/h
t = set_vfrac(t, 'E-601', '100.', '1.0', '0.174')
t = set_vfrac(t, 'E-602', '80.', '0.47', '0.210')
t = set_vfrac(t, 'E-603', '60.', '0.20', '0.266')
print('3) 三效蒸发 汽化率 0.174 / 0.210 / 0.266（SPEC-OPT=PV）✓')

# ---------- 4) 结晶器：常压冷却到 25 ℃ ----------
t = set_flash(t, 'C-601', '25.', '1.0')
print('4) C-601 结晶器 -> 25 ℃ / 1.0 bar ✓')

# ---------- 5) 脱色（M-502）分离分数：只脱色体与少量夹带液 ----------
for cid, v in [('NAC', 0.9), ('3-CP', 0.05), ('NAM', 0.005), ('H2O', 0.005),
               ('3-MP', 0.0), ('NH3', 0.0), ('O2', 0.0), ('N2', 0.0), ('4-CP', 0.0),
               ('TOL', 0.0), ('CO2', 0.0), ('HCN', 0.0), ('CO', 0.0), ('AIR', 0.0)]:
    t = set_frac(t, 'M-502', cid, v)
print('5) M-502 脱色分离分数 ✓（NAC 0.9 / 3-CP 0.05 / NAM 0.005）')

# ---------- 6) 离心机（M-601）分离分数 ----------
for cid, v in [('NAM', 0.05), ('H2O', 0.90), ('NAC', 0.05), ('3-CP', 0.95),
               ('3-MP', 0.95), ('NH3', 0.0), ('O2', 0.0), ('N2', 0.0), ('4-CP', 0.0),
               ('TOL', 0.0), ('CO2', 0.0), ('HCN', 0.0), ('CO', 0.0), ('AIR', 0.0)]:
    t = set_frac(t, 'M-601', cid, v)
print('6) M-601 离心机分离分数 ✓（NAM 0.05 进母液）')

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
while time.time() - t0 < 480:
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
L.append('=== 面板消息 (%d) ===' % len(MSGS))
for i2, x in enumerate(MSGS):
    if re.search(r'ERROR|SEVERE|WARNING|Terminal|Severe|Errors|Warnings|completed|BYPASSED|CONVERG', x):
        L.append('%3d | %s' % (i2, x[:170]))

L.append('')
L.append('=== 全部流股（质量流量 kg/h）===')
ALL = ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-132', 'S-210',
       'S-104', 'S-105', 'S-106', 'S-129', 'S-108', 'S-109', 'S-110', 'S-112',
       'S-113', 'S-114', 'S-114A', 'S-115', 'S-116', 'S-116A', 'S-117', 'S-118',
       'S-119', 'S-120', 'S-121', 'S-122', 'S-130', 'S-131', 'S-126', 'S-127', 'S-128',
       'S-201', 'S-202', 'S-203', 'S-211', 'S-212', 'S-213',
       'S-401', 'S-402', 'S-214', 'S-205', 'S-403', 'S-206', 'S-404', 'S-207',
       'S-405', 'S-208', 'S-406', 'S-209']
for s in ALL:
    L.append('  %-8s %-12s' % (s, g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s)))

L.append('')
L.append('=== 总物料平衡 ===')
IN = ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-132', 'S-210']
OUTL = ['S-130', 'S-131', 'S-211', 'S-212', 'S-213', 'S-115', 'S-117', 'S-119',
        'S-122', 'S-405', 'S-406', 'S-209', 'S-402', 'S-205', 'S-403']


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
L.append('  进料合计 = %.2f kg/h   缺:%s' % (vi, mi))
L.append('  出料合计 = %.2f kg/h   缺:%s' % (vo, mo))
L.append('  偏差 = %.2f kg/h  (%.3f %%)' % (vo - vi, (vo - vi) / vi * 100 if vi else 0))

L.append('')
L.append('=== 关键流股组成 ===')
for s in ['S-203', 'S-401', 'S-214', 'S-403', 'S-404', 'S-208', 'S-209', 'S-405', 'S-212', 'S-213']:
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

open(r'D:\<化工工作区>\_probe\s6.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('DONE')
