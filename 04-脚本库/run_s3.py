# -*- coding: utf-8 -*-
"""S3：修正 E-602 进料 / 三效蒸发温度 / T-302 汽提塔收敛"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_s2.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_s3.bkp'
t = open(SRC, encoding='utf-8', errors='ignore').read()
print('源大小', len(t))


def slice_entry(x, blkid):
    i = x.find('BLKID = "%s"' % blkid)
    assert i > 0, blkid
    j = x.find('BLOCK', i + 20)
    if j < 0:
        j = x.find('"DEF-STREAM"', i + 20)
    return i, j


def set_in(x, blkid, newin):
    i, j = slice_entry(x, blkid)
    seg = x[i:j]
    seg2 = re.sub(r'IN = \([^)]*\)', 'IN = ( %s )' % newin, seg, count=1)
    assert seg2 != seg, blkid
    return x[:i] + seg2 + x[j:]


def repl_block(x, bid, fn):
    starts = [s.start() for s in re.finditer(r'\?\s*BLOCK\s+', x)]
    for k in range(len(starts)):
        s0 = starts[k]
        s1 = starts[k + 1] if k + 1 < len(starts) else len(x)
        m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', x[s0:s1])
        if m and m.group(2) == bid:
            return x[:s0] + fn(x[s0:s1]) + x[s1:]
    raise RuntimeError(bid)


# 1) E-602 进料改接 E-601 的液相出口 S-401
t = set_in(t, 'E-602', '"S-401" M0-1')
print('1) E-602 进料 -> S-401（E-601 液相）✓')

# 2) 三效蒸发温度/压力
for bid, temp, pres in [('E-601', '100.', '1.0'), ('E-602', '80.', '0.47'), ('E-603', '60.', '0.20')]:
    def fn(seg, bid=bid, temp=temp, pres=pres):
        seg = re.sub(r'TEMP = [\d.]+ <22> <4>', 'TEMP = %s <22> <4>' % temp, seg, count=1)
        seg = re.sub(r'PRES = [\d.]+ <20> <5>', 'PRES = %s <20> <5>' % pres, seg, count=1)
        return seg
    t = repl_block(t, bid, fn)
    print('   %s -> %s ℃ / %s bar' % (bid, temp, pres))
print('2) 三效蒸发参数 ✓')

# 3) T-302 汽提塔：降板数、提高再沸比
def fn302(seg):
    seg = seg.replace('PARAM NSTAGE = 8 NSTAGEMAX = 9', 'PARAM NSTAGE = 5 NSTAGEMAX = 6')
    seg = seg.replace('PROD-STAGE = 8 PROD-PHASE = L', 'PROD-STAGE = 5 PROD-PHASE = L')
    seg = seg.replace('TEMP-STAGE = 8 TEMP-EST = 100.0', 'TEMP-STAGE = 5 TEMP-EST = 100.0')
    seg = seg.replace('BASIS-BR = 0.1 <-1> <0>', 'BASIS-BR = 0.5 <-1> <0>')
    return seg


t = repl_block(t, 'T-302', fn302)
print('3) T-302 -> 5 板 / 再沸比 0.5 ✓')

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
for _ in range(60):
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


print()
print('=== 错误与汇总 ===')
for x in MSGS:
    if re.search(r'ERROR|SEVERE|NOT CONVERGED|Summary|Terminal|Warnings.*\d|completed', x):
        print('   |', x[:165])
print()
print('=== 后处理段 ===')
for s in ['S-201', 'S-203', 'S-401', 'S-402', 'S-214', 'S-205', 'S-403', 'S-206',
          'S-404', 'S-207', 'S-405', 'S-208', 'S-406', 'S-209', 'S-212', 'S-213', 'S-211']:
    print('  %-7s MASS=%-11s' % (s, g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s)))
print()
print('=== 主流程关键流股（校核是否不变）===')
for s in ['S-104', 'S-106', 'S-110', 'S-113', 'S-114', 'S-115', 'S-116', 'S-118',
          'S-119', 'S-120', 'S-121', 'S-122']:
    print('  %-7s MASS=%-11s' % (s, g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s)))
print()
print('=== 产品 S-209（干品）组成 ===')
nn = doc.Tree.FindNode(r'\Data\Streams\S-209\Output\MASSFLOW3')
if nn is not None:
    for i2 in range(nn.Elements.Count):
        e = nn.Elements.Item(i2)
        try:
            if e.Value and abs(float(e.Value)) > 0.01:
                print('     %-8s %9.2f' % (e.Name, float(e.Value)))
        except Exception:
            pass
try:
    doc.SaveAs(OUT.replace('.bkp', '_s.bkp'))
    print('已保存')
except Exception as ex:
    print('保存失败:', ex)
try:
    doc.Close()
except Exception:
    pass
print('DONE')
