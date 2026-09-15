# -*- coding: utf-8 -*-
"""v8：先跑一次读各塔进料组成 -> 按组分分配算出正确的 D:F -> 再跑一次验证塔温梯度"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_v7.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_v8.bkp'

MW = {'3-MP': 93.13, '4-MP': 93.13, 'NH3': 17.03, 'O2': 32.00, 'N2': 28.01,
      'H2O': 18.02, '3-CP': 104.11, '4-CP': 104.11, 'TOL': 92.14, 'NAM': 122.13,
      'NAC': 123.11, 'NAOH': 40.00, 'H2SO4': 98.08, 'NA2SO4': 142.04,
      'CO2': 44.01, 'CO': 28.01, 'HCN': 27.03, 'AIR': 28.96}


class Runner:
    def __init__(self):
        self.msgs = []

    def run(self, path, tag=''):
        msgs = self.msgs
        msgs.clear()

        class Sink:
            def OnControlPanelMessage(self, *a):
                s = ' '.join(str(x) for x in a).strip()
                if s and s != 'False':
                    msgs.append(s)

        doc = win32.DispatchEx('Apwn.Document')
        try:
            doc.SuppressDialogs = True
        except Exception:
            pass
        try:
            win32.WithEvents(doc, Sink)
        except Exception:
            pass
        doc.InitFromArchive2(path)
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
        return doc


def g(doc, p):
    n = doc.Tree.FindNode(p)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return None


def comps(doc, s):
    """返回 {组分: 质量流量 kg/h}, 总摩尔流 kmol/h"""
    b = r'\Data\Streams\%s\Output' % s
    n = doc.Tree.FindNode(b + r'\MASSFLOW3')
    d = {}
    if n is not None:
        for i in range(n.Elements.Count):
            try:
                e = n.Elements.Item(i)
                v = e.Value
                if v:
                    d[e.Name] = float(v)
            except Exception:
                pass
    tot = g(doc, b + r'\MOLEFLMX')
    return d, tot


R = Runner()
print('===== 第 1 次运行：读组成 =====', flush=True)
doc = R.run(SRC, 'read')
for x in R.msgs:
    if 'ERROR' in x or 'MASS BALANCE' in x or 'completed' in x:
        print('   |', x[:170], flush=True)

print()
print('===== 各流股组成 (kg/h) =====', flush=True)
STREAMS = ['S-110', 'S-112', 'S-113', 'S-114', 'S-115', 'S-116', 'S-117', 'S-118',
           'S-119', 'S-120', 'S-121', 'S-122']
data = {}
for s in STREAMS:
    d, tot = comps(doc, s)
    data[s] = (d, tot)
    items = ', '.join('%s=%.2f' % (k, v) for k, v in sorted(d.items(), key=lambda kv: -kv[1]) if v > 1e-6)
    print('  %-7s tot=%s kmol/h | %s' % (s, ('%.3f' % tot) if tot else '?', items[:230]), flush=True)

print()
print('===== 摩尔分率与建议 D:F =====', flush=True)


def zmo(d):
    tot = 0.0
    for k, v in d.items():
        tot += v / MW.get(k, 100.0)
    return {k: (v / MW.get(k, 100.0)) / tot for k, v in d.items()}, tot


rec = {}
for s in STREAMS:
    d, _ = data[s]
    if not d:
        continue
    z, _ = zmo(d)
    print('  %-7s %s' % (s, ', '.join('%s=%.4f' % (k, v) for k, v in
                                      sorted(z.items(), key=lambda kv: -kv[1])[:6])), flush=True)
    rec[s] = z

print()
# T-401: 顶出甲苯   D:F = z_TOL(进料 S-114)
d401 = rec.get('S-114', {})
print('  T-401 建议 D:F = z(TOL) = %.4f' % d401.get('TOL', 0.0), flush=True)
# T-402: 顶出水     D:F = z_H2O(进料 S-116)
d402 = rec.get('S-116', {})
print('  T-402 建议 D:F = z(H2O) = %.4f' % d402.get('H2O', 0.0), flush=True)
# T-403: 顶出 4-CP  D:F = z(4-CP)(进料 S-118)，留一点余量
d403 = rec.get('S-118', {})
print('  T-403 建议 D:F = z(4-CP) = %.5f' % d403.get('4-CP', 0.0), flush=True)
# T-404: 顶出 3-CP 产品 D:F = z(3-CP)(进料 S-120)
d404 = rec.get('S-120', {})
print('  T-404 建议 D:F = z(3-CP) = %.4f' % d404.get('3-CP', 0.0), flush=True)

try:
    doc.Close()
except Exception:
    pass
time.sleep(1)

# ---------- 打补丁 ----------
NEWDF = {
    'T-401': 0.0,   # 占位，下面按计算填
    'T-402': d402.get('H2O', 0.5),
    'T-403': max(d403.get('4-CP', 0.01) * 1.10, 0.002),
    'T-404': 0.98,
}
NEWDF['T-401'] = d401.get('TOL', 0.5)

t = open(SRC, encoding='utf-8', errors='ignore').read()
print()
print('===== 写入新 D:F =====', flush=True)
starts = [m.start() for m in re.finditer(r'\?\s*BLOCK\s+[A-Z0-9]+\s+"?[A-Za-z0-9-]+"?\s*\?', t)]
for idx in range(len(starts) - 1, -1, -1):
    s0 = starts[idx]
    s1 = starts[idx + 1] if idx + 1 < len(starts) else len(t)
    m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"([^"]+)"\s*\?', t[s0:s1])
    if not m:
        continue
    bid = m.group(2)
    if bid in NEWDF:
        seg = t[s0:s1]
        v = NEWDF[bid]
        seg2 = re.sub(r'D:F = [\d.]+ <-1> <0>', 'D:F = %.5f <-1> <0>' % v, seg, count=1)
        if seg2 != seg:
            t = t[:s0] + seg2 + t[s1:]
            print('   %s  D:F -> %.5f' % (bid, v), flush=True)
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)

print()
print('===== 第 2 次运行：验证塔温梯度 =====', flush=True)
R2 = Runner()
doc = R2.run(OUT, 'verify')
for x in R2.msgs:
    if 'ERROR' in x or 'SEVERE' in x or 'Converged' in x or 'completed' in x:
        print('   |', x[:170], flush=True)

print()
print('===== 塔结果 =====', flush=True)
for b in ['T-201', 'T-301', 'T-401', 'T-402', 'T-403', 'T-404']:
    bb = r'\Data\Blocks\%s\Output' % b
    print('  %-7s COND=%s REB=%s RR=%s Ttop=%s Tbot=%s' % (
        b, g(doc, bb + r'\COND_DUTY'), g(doc, bb + r'\REB_DUTY'), g(doc, bb + r'\MOLE_RR'),
        g(doc, bb + r'\TOP_TEMP'), g(doc, bb + r'\BOTTOM_TEMP')), flush=True)

print()
print('===== 关键流股 =====', flush=True)
for s in ['S-114', 'S-115', 'S-116', 'S-117', 'S-118', 'S-119', 'S-120', 'S-121', 'S-122']:
    b = r'\Data\Streams\%s\Output' % s
    d, tot = comps(doc, s)
    top = ', '.join('%s=%.1f' % (k, v) for k, v in sorted(d.items(), key=lambda kv: -kv[1])[:4] if v > 1e-6)
    print('  %-7s T=%s P=%s N=%.3f W=%.1f | %s' % (
        s, g(doc, b + r'\TEMP_OUT'), g(doc, b + r'\PRES_OUT'),
        tot or 0, g(doc, b + r'\MASSFLMX') or 0, top), flush=True)

for p in [OUT, OUT.replace('.bkp', '.apwz')]:
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
