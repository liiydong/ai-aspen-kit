# -*- coding: utf-8 -*-
"""定稿模型全流程复核：拓扑 + 控制面板 + 流股结果 + 塔水力学"""
import sys, re, time, os, json
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_最终定稿.bkp'
t = open(SRC, encoding='utf-8', errors='ignore').read()
L = []
L.append('文件: %s   %d 字节' % (os.path.basename(SRC), len(t)))

# ---------- 1) 从文本解析拓扑 ----------
def text_flowsheet(x):
    m = re.search(r'\?\s*FLOWSHEET\s+GLOBAL\s*\?', x)
    if not m:
        return None
    s0 = m.end()
    m2 = re.search(r'\n\?\s', x[s0:])
    s1 = s0 + (m2.start() if m2 else len(x) - s0)
    return x[s0:s1]

fs = text_flowsheet(t)
conns = []
if fs:
    # 归一化：去换行、折叠空格
    flat = re.sub(r'\s+', ' ', fs)
    # 每条: BLOCK BLKID = "name" SECTION = n IN = ( ... ) OUT = ( ... )
    for m in re.finditer(r'BLOCK\s+BLKID\s*=\s*"?([A-Za-z0-9_-]+)"?\s*(?:SECTION\s*=\s*\d+\s*)?IN\s*=\s*\(\s*([^)]*)\)\s*OUT\s*=\s*\(\s*([^)]*)\)', flat):
        bid = m.group(1)
        ins = re.findall(r'"([^"]+)"', m.group(2))
        outs = re.findall(r'"([^"]+)"', m.group(3))
        conns.append((bid, ins, outs))
L.append('FLOWSHEET 解析: %d 条块连接' % len(conns))

# 流股名全集
allstreams = set()
for bid, ins, outs in conns:
    allstreams |= set(ins) | set(outs)
# 流股段定义
sec_streams = set(re.findall(r'\?\s*STREAM\s+MATERIAL\s+"([^"]+)"\s*\?', re.sub(r'\n', ' ', t)))
sec_streams |= set(re.findall(r'\?\s*STREAM\s+MATERIAL\s*\n\s*"([^"]+)"\s*\?', t))
L.append('FLOWSHEET 引用流股: %d 条；STREAM 段定义: %d 条' % (len(allstreams), len(sec_streams)))

# 块实例注册表（文件头）
reg = re.findall(r'BLOCK\s+ID\s*=\s*"?([A-Za-z0-9_-]+)"?\s+TYPE\s*=\s*([A-Za-z0-9]+)', t)
registry = dict(reg)

# ---------- 2) 打开并运行 ----------
MSGS = []
class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            MSGS.append(s)

doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
win32.WithEvents(doc, Sink)
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

def comp(s, mode='MASSFLOW3'):
    nn = doc.Tree.FindNode(r'\Data\Streams\%s\Output\%s' % (s, mode))
    d = {}
    if nn is None:
        return d
    try:
        for i3 in range(nn.Elements.Count):
            e = nn.Elements.Item(i3)
            try:
                v = float(e.Value) if e.Value is not None else 0.0
                if abs(v) > 1e-6:
                    d[e.Name] = v
            except Exception:
                pass
    except Exception:
        pass
    return d

# ---------- 3) 结果 ----------
L.append('')
L.append('======== 控制面板 ========')
for i2, x in enumerate(MSGS):
    L.append('%3d | %s' % (i2, x[:190]))

# 块清单
L.append('')
L.append('======== 块清单 (COM) ========')
blknode = doc.Tree.FindNode(r'\Data\Blocks')
bids = []
if blknode is not None:
    for i in range(blknode.Elements.Count):
        e = blknode.Elements.Item(i)
        nm = e.Name
        bids.append(nm)
        mt = g(r'\Data\Blocks\%s\Input\MODEL' % nm)
        L.append('  %-8s MODEL=%-12s TYPE=%s' % (nm, mt, registry.get(nm, '?')))
L.append('  COM 读到块数 = %d ; 文本注册表 = %d' % (len(bids), len(registry)))

# 流股结果
L.append('')
L.append('======== 流股结果 ========')
src_map = {}
dst_map = {}
for bid, ins, outs in conns:
    for s in ins:
        dst_map.setdefault(s, []).append(bid)
    for s in outs:
        src_map.setdefault(s, []).append(bid)

sn = doc.Tree.FindNode(r'\Data\Streams')
sids = []
if sn is not None:
    for i in range(sn.Elements.Count):
        sids.append(sn.Elements.Item(i).Name)
L.append('  COM 读到流股数 = %d' % len(sids))
L.append('  %-9s %10s %9s %9s  %-22s %-22s' % ('流股', 'kg/h', 'T(℃)', 'P(bar)', '来源', '去向'))
zero = []
for s in sids:
    fl = g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s)
    T = g(r'\Data\Streams\%s\Output\TEMP_OUT\MIXED' % s)
    P = g(r'\Data\Streams\%s\Output\PRES_OUT\MIXED' % s)
    up = ','.join(src_map.get(s, [])) or '(进料)'
    dn = ','.join(dst_map.get(s, [])) or '(出料)'
    L.append('  %-9s %10s %9s %9s  %-22s %-22s' % (
        s, ('%.2f' % fl) if fl is not None else 'None',
        ('%.2f' % (T - 273.15)) if T not in (None, 0) else '-',
        ('%.4f' % P) if P not in (None, 0) else '-', up, dn))
    if fl is None or abs(float(fl)) < 1e-6:
        zero.append(s)
L.append('  零/空流量流股: %s' % (', '.join(zero) if zero else '无'))

# ---------- 4) 物料平衡 ----------
IN = ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-132', 'S-210']
OUTL = ['S-130', 'S-131', 'S-211', 'S-212', 'S-213', 'S-115', 'S-119', 'S-122',
        'S-216', 'S-217', 'S-218', 'S-405C', 'S-406', 'S-209']
vi = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in IN)
vo = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in OUTL)
L.append('')
L.append('======== 总物料平衡 ========')
L.append('  进料 %s' % ' + '.join('%s=%.2f' % (s, float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0)) for s in IN))
L.append('  出料 %s' % ' + '.join('%s=%.2f' % (s, float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0)) for s in OUTL))
L.append('  进 %.2f  出 %.2f  偏差 %.4f %%' % (vi, vo, (vo - vi) / vi * 100 if vi else 0))

# ---------- 5) 塔参数与水力学 ----------
L.append('')
L.append('======== 塔/关键单元参数 ========')
TOWERS = ['T-201', 'T-202', 'T-203', 'T-301', 'T-302', 'T-401', 'T-402', 'T-403', 'T-404']
for b in TOWERS:
    ns = g(r'\Data\Blocks\%s\Input\NSTAGE' % b)
    rr = g(r'\Data\Blocks\%s\Output\RR' % b)
    cd = g(r'\Data\Blocks\%s\Output\COND_DUTY' % b)
    rd = g(r'\Data\Blocks\%s\Output\REB_DUTY' % b)
    tt = g(r'\Data\Blocks\%s\Output\TOPD_TEMP' % b)
    bt = g(r'\Data\Blocks\%s\Output\BOTM_TEMP' % b)
    df = g(r'\Data\Blocks\%s\Output\D:F' % b)
    def f(v, k=1):
        try:
            return '%.4g' % (float(v) * k)
        except Exception:
            return str(v)
    L.append('  %-7s NSTAGE=%-5s RR=%-8s D:F=%-8s 顶=%-8s 底=%-8s QC=%-12s QR=%s'
             % (b, ns, f(rr), f(df), f(tt), f(bt), f(cd), f(rd)))

# ---------- 6) 产品与收率 ----------
L.append('')
L.append('======== 产品与收率 ========')
p = comp('S-209')
tt = sum(p.values())
L.append('  S-209 = %.2f kg/h' % tt)
for k, v in sorted(p.items(), key=lambda x: -x[1]):
    L.append('     %-8s %9.3f  %6.3f wt%%' % (k, v, v / tt * 100 if tt else 0))
L.append('  折年产(7200h) = %.0f t/a' % (tt / 1000 * 7200))
nm = p.get('NAM', 0)
mp3 = comp('S-101').get('3-MP', 0)
L.append('  3-MP 单耗 = %.1f kg/t' % (mp3 / nm * 1000 if nm else 0))
L.append('  总收率 = %.2f%%' % (nm / 122.13 / (mp3 / 93.13) * 100 if mp3 else 0))

# ---------- 7) 水力学 ----------
L.append('')
L.append('======== 塔水力学（供第5章塔径校核）========')
for b in TOWERS:
    vf = g(r'\Data\Blocks\%s\Output\VAP_VFLOW' % b)
    lf = g(r'\Data\Blocks\%s\Output\LIQ_VFLOW' % b)
    vd = g(r'\Data\Blocks\%s\Output\VAP_DENS' % b)
    ld = g(r'\Data\Blocks\%s\Output\LIQ_DENS' % b)
    L.append('  %-7s VAP_VFLOW=%-12s LIQ_VFLOW=%-12s VAP_DENS=%-10s LIQ_DENS=%s'
             % (b, vf, lf, vd, ld))

# 保存完整结果
open(r'D:\<化工工作区>\_probe\verify_all.txt', 'w', encoding='utf-8').write('\n'.join(L))

# 导出报告文件
try:
    doc.Export(2, r'D:\<化工工作区>\_probe\final_report.txt')
    doc.Export(6, r'D:\<化工工作区>\_probe\final_panel.txt')
except Exception as ex:
    print('export fail', ex)
try:
    doc.SaveAs(r'D:\<化工工作区>\NA-Chemical-10000t_提交版.bkp')
except Exception as ex:
    print('save fail', ex)
try:
    doc.Close()
except Exception:
    pass
print('DONE  lines=%d' % len(L))
