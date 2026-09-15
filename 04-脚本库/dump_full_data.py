# -*- coding: utf-8 -*-
"""采集全流程完整数据（供论文改写）"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_最终定稿.bkp'
t = open(SRC, encoding='utf-8', errors='ignore').read()
L = []

# 拓扑与类型
m = re.search(r'\?\s*FLOWSHEET\s+GLOBAL\s*\?', t)
s0 = m.end()
m2 = re.search(r'\n\?\s*[A-Z]', t[s0:])
fs = re.sub(r'\s+', ' ', t[s0:s0 + (m2.start() if m2 else 20000)])
conns = {}
for r in re.split(r'(?=BLOCK\s+BLKID\s*=)', fs):
    mm = re.match(r'BLOCK\s+BLKID\s*=\s*"([^"]+)"\s+BLKTYPE\s*=\s*"([^"]+)"', r)
    if not mm:
        continue
    ins = re.findall(r'"([^"]+)"', (re.search(r'IN\s*=\s*\((.*?)\)', r) or [None, ''])[1])
    outs = re.findall(r'"([^"]+)"', (re.search(r'OUT\s*=\s*\((.*?)\)', r) or [None, ''])[1])
    conns[mm.group(1)] = (mm.group(2), ins, outs)
srcmap, dstmap = {}, {}
for b, (bt, ins, outs) in conns.items():
    for s in ins:
        dstmap.setdefault(s, []).append(b)
    for s in outs:
        srcmap.setdefault(s, []).append(b)

doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
doc.InitFromArchive2(SRC)
time.sleep(3)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 900:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 10:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(100):
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

sn = doc.Tree.FindNode(r'\Data\Streams')
sids = [sn.Elements.Item(i).Name for i in range(sn.Elements.Count)]

L.append('=========== 全流股数据（Aspen 定稿模型）===========')
L.append('%-9s %9s %9s %9s %8s | %-24s | %-26s | %s' % ('流股', 'kg/h', 'kmol/h', 'T ℃', 'P bar', '来源', '去向', '组成 wt%'))
for s in sids:
    fl = g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s)
    try:
        fl = float(fl)
    except Exception:
        fl = None
    mo = g(r'\Data\Streams\%s\Output\MOLEFLMX\MIXED' % s)
    try:
        mo = float(mo)
    except Exception:
        mo = None
    T = g(r'\Data\Streams\%s\Output\TEMP_OUT\MIXED' % s)
    P = g(r'\Data\Streams\%s\Output\PRES_OUT\MIXED' % s)
    c = comp(s)
    tt = sum(c.values())
    cstr = ', '.join('%s %.2f' % (k, v / tt * 100) for k, v in sorted(c.items(), key=lambda x: -x[1])[:6]) if tt else ''
    L.append('%-9s %9s %9s %9s %8s | %-24s | %-26s | %s' % (
        s, ('%.2f' % fl) if fl is not None else '-', ('%.3f' % mo) if mo is not None else '-',
        ('%.2f' % T) if T is not None else '-', ('%.4f' % P) if P is not None else '-',
        ','.join(srcmap.get(s, [])) or '(外部进料)', ','.join(dstmap.get(s, [])) or '(出料)', cstr))

L.append('')
L.append('=========== 关键流股详细组成 ===========')
for s in ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-210', 'S-132',
          'S-105', 'S-106', 'S-109', 'S-110', 'S-113', 'S-114', 'S-115', 'S-116',
          'S-117', 'S-118', 'S-119', 'S-120', 'S-121', 'S-122', 'S-130', 'S-131',
          'S-201', 'S-202', 'S-203', 'S-205', 'S-206', 'S-211', 'S-212', 'S-213',
          'S-401', 'S-402', 'S-403', 'S-404', 'S-405', 'S-405B', 'S-405C', 'S-208',
          'S-209', 'S-216', 'S-217', 'S-218', 'S-406']:
    c = comp(s)
    tt = sum(c.values())
    L.append('  %-8s %9.2f kg/h  T=%s P=%s' % (s, tt, g(r'\Data\Streams\%s\Output\TEMP_OUT\MIXED' % s), g(r'\Data\Streams\%s\Output\PRES_OUT\MIXED' % s)))
    for k, v in sorted(c.items(), key=lambda x: -x[1]):
        L.append('        %-8s %10.3f  %7.3f wt%%' % (k, v, v / tt * 100 if tt else 0))

L.append('')
L.append('=========== 全部块参数 ===========')
for b in sorted(conns):
    bt = conns[b][0]
    items = []
    for k in ['NSTAGE', 'BASIS_RR', 'D:F', 'PRES1', 'TEMP', 'PRES', 'CONDENSER', 'REBOILER',
              'SPEC-OPT', 'VFRAC', 'FRACS', 'TEMP_OPT', 'T_EST', 'P_SPEC', 'FRAC',
              'DELP', 'PRES2', 'PARAM']:
        v = g(r'\Data\Blocks\%s\Input\%s' % (b, k))
        if v is not None:
            items.append('%s=%s' % (k, str(v)[:20]))
    for k in ['TOP_TEMP', 'BOTTOM_TEMP', 'COND_DUTY', 'REB_DUTY', 'RR']:
        v = g(r'\Data\Blocks\%s\Output\%s' % (b, k))
        if v is not None:
            try:
                if abs(float(v)) > 1e-9:
                    items.append('%s=%s' % (k, ('%.4g' % float(v))))
            except Exception:
                items.append('%s=%s' % (k, str(v)[:20]))
    L.append('  %-8s %-9s IN=[%s] OUT=[%s] %s' % (b, bt, ','.join(conns[b][1]), ','.join(conns[b][2]), ' '.join(items)))

# 物料平衡
IN = ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-132', 'S-210']
OUTL = ['S-130', 'S-131', 'S-211', 'S-212', 'S-213', 'S-115', 'S-119', 'S-122',
        'S-216', 'S-217', 'S-218', 'S-405C', 'S-406', 'S-209']
vi = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in IN)
vo = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in OUTL)
L.append('')
L.append('=========== 物料平衡 ===========')
L.append('  进 %.2f  kg/h  = 出 %.2f kg/h   偏差 %.5f%%' % (vi, vo, (vo - vi) / vi * 100))
L.append('  进料明细: %s' % ', '.join('%s=%.2f' % (s, float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0)) for s in IN))
L.append('  出料明细: %s' % ', '.join('%s=%.2f' % (s, float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0)) for s in OUTL))
p = comp('S-209'); tt = sum(p.values())
L.append('  产品 S-209 %.2f kg/h  折年产 %.0f t/a' % (tt, tt / 1000 * 7200))
nm = p.get('NAM', 0); mp3 = comp('S-101').get('3-MP', 0)
L.append('  NAM %.2f kg/h (%.3f%%)  3-MP 单耗 %.1f kg/t  总收率 %.2f%%'
         % (nm, nm / tt * 100 if tt else 0, mp3 / nm * 1000 if nm else 0,
            nm / 122.13 / (mp3 / 93.13) * 100 if mp3 else 0))

open(r'D:\<化工工作区>\_probe\full_data.txt', 'w', encoding='utf-8').write('\n'.join(L))
try:
    doc.Export(2, r'D:\<化工工作区>\_probe\report_final.txt')
except Exception:
    pass
try:
    doc.Close()
except Exception:
    pass
print('DONE lines=%d' % len(L))
