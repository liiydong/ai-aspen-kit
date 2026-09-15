# -*- coding: utf-8 -*-
"""P7（最终重建）：从干净的 bipA2 重建，所有改动一次到位并回读校验"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\_probe\bipA2.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_v20.bkp'
t = open(SRC, encoding='utf-8', errors='ignore').read()
assert len(re.findall(r'\?\s*BLOCK\s+RADFRAC', t)) == 5, '源文件不干净'
print('源文件干净检查通过，大小', len(t))


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


def repl_block(t, bid, fn):
    """对指定块段落执行 fn，只替换第一处"""
    starts = [s.start() for s in re.finditer(r'\?\s*BLOCK\s+', t)]
    for k in range(len(starts)):
        s0 = starts[k]
        s1 = starts[k + 1] if k + 1 < len(starts) else len(t)
        m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
        if m and m.group(2) == bid:
            return t[:s0] + fn(t[s0:s1]) + t[s1:]
    raise RuntimeError('未找到块 ' + bid)


# ---- a) 注册表改类型 ----
t2 = t.replace('T-301\nExtract\nBuilt-In\nEXTRACT\n', 'T-301\nSep\nBuilt-In\nSEP\n', 1)
assert t2 != t, '注册表未改'
t = t2
print('a) 注册表 -> Sep ✓')

# ---- b) FLOWSHEET：SEP + 端口 M0-1 + 断开 S-115 ----
old_fs = ('BLKID = "T-301" BLKTYPE = "EXTRACT" MDLTYPE = "Extract" IN = ( "S-112" M0-1 '
          '\n"S-111" M1-2 "S-115" M2-3 ) OUT = ( "S-113" M1-2 "S-114" M0-1 )')
new_fs = ('BLKID = "T-301" BLKTYPE = "SEP" MDLTYPE = "Sep" IN = ( "S-112" M0-1 '
          '"S-111" M0-1 ) OUT = ( "S-113" M0-1 "S-114" M0-1 )')
assert old_fs in t, 'FLOWSHEET 未匹配'
t = t.replace(old_fs, new_fs, 1)
print('b) FLOWSHEET -> Sep，断开 S-115 循环 ✓')

# ---- c) Sep 参数段 ----
FR = {'H2O': 0.999, 'TOL': 0.0005, '3-CP': 0.005, '4-CP': 0.005,
      '3-MP': 0.005, '4-MP': 0.005, 'HCN': 0.50, 'NH3': 0.99,
      'O2': 0.99, 'N2': 0.99, 'CO2': 0.99, 'CO': 0.99,
      'NAM': 0.999, 'NAC': 0.999, 'AIR': 0.999}
ci = t.find('? COMPONENTS MAIN ?')
cj = t.find('? COMPONENTS "COMP-LIST"', ci)
cids = re.findall(r'CID = "?([A-Za-z0-9-]+)"?', t[ci:cj])
recs = ['PARAM-STREAM = "S-113" SUBSTREAM = MIXED COMPS = "%s" FRACS = %s <0> <0> ' % (c, FR[c])
        for c in cids if c in FR]
sep_para = ('? BLOCK SEP "T-301" ? ; "METCBAR_MOLE" ; ; ICON2 ; \\ '
            'DESCRIPTION DESCRIPTION = "Toluene extraction 99.5 pct per Ruibang" \\ \\ '
            'PARAM1 PRES1 = 1.0 <20> <5> \\ \\ PARAM ' + ' /  '.join(recs) + ' \\ ')
t = repl_block(t, 'T-301', lambda seg: wr(sep_para))
print('c) T-301 -> SEP 段 ✓（组分 %d 个）' % len(recs))

# ---- d) 塔压 ----
for bid, pres in [('T-401', '0.24'), ('T-402', '0.26'), ('T-403', '0.40'), ('T-404', '0.30')]:
    def fn(seg, p=pres):
        s = re.sub(r'PRES1 = [\d.]+ <20> <5>', 'PRES1 = %s <20> <5>' % p, seg, count=1)
        return s
    t = repl_block(t, bid, fn)
print('d) 四塔压力 -> 0.24/0.26/0.40/0.30 bar ✓')


# ---- e) T-401 采出比与回流比 ----
def fn401(seg):
    seg = re.sub(r'D:F = [\d.]+ <-1> <0>', 'D:F = 0.6590 <-1> <0>', seg, count=1)
    seg = re.sub(r'BASIS-RR = [\d.]+ <-1> <0>', 'BASIS-RR = 3.0 <-1> <0>', seg, count=1)
    return seg


t = repl_block(t, 'T-401', fn401)
print('e) T-401 D:F=0.6590 RR=3.0 ✓')

# ---- f) S-111 = 设计甲苯循环量 ----
m = re.search(r'\?\s*STREAM\s+MATERIAL\s+"S-111"\s*\?', t)
assert m, '未找到 S-111 段'
i = m.start()
nxt = re.search(r'\?\s*STREAM\s+MATERIAL\s+"S-1', t[i + 30:])
j = i + 30 + nxt.start()
seg = t[i:j]
seg2 = re.sub(r'(TOTAL\s*=\s*)[\d.]+', r'\g<1>24.47', seg)
seg2 = re.sub(r'(CID\s*=\s*TOL\s+FLOW\s*=\s*)[\d.]+', r'\g<1>24.47', seg2)
seg2 = re.sub(r'(TEMP\s*=\s*)[\d.]+', r'\g<1>40', seg2)
t = t[:i] + seg2 + t[j:]
print('f) S-111 = 24.47 kmol/h 甲苯（设计循环量）✓')

# ---- 重复性校验 ----
assert len(re.findall(r'\?\s*BLOCK\s+RADFRAC', t)) == 5
assert len(re.findall(r'BLOCK\s+RADFRAC\s*\n?\s*"T-401"', t)) == 1
assert len(re.findall(r'\?\s*STREAM\s+MATERIAL\s*\n?\s*"S-111"', t)) == 1
print('重复性校验通过（RADFRAC=5, T-401=1, S-111=1）✓')
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)
print('写出:', OUT, os.path.getsize(OUT))

# ---- 运行 ----
MSGS = []


class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            MSGS.append(s)


doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
try:
    win32.WithEvents(doc, Sink)
except Exception:
    pass
doc.InitFromArchive2(OUT)
time.sleep(3)


def g(p):
    n = doc.Tree.FindNode(p)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return None


print()
print('=== 参数回读校验 ===')
print('  T-401 D:F    =', g(r'\Data\Blocks\T-401\Input\D:F'))
print('  T-401 BASIS_RR =', g(r'\Data\Blocks\T-401\Input\BASIS_RR'))
print('  T-401 PRES1  =', g(r'\Data\Blocks\T-401\Input\PRES1'))
print('  T-402 PRES1  =', g(r'\Data\Blocks\T-402\Input\PRES1'))
print('  T-403 PRES1  =', g(r'\Data\Blocks\T-403\Input\PRES1'))
print('  T-404 PRES1  =', g(r'\Data\Blocks\T-404\Input\PRES1'))

doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 400:
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


def c(s, cid):
    n = doc.Tree.FindNode(r'\Data\Streams\%s\Output\MASSFLOW3' % s)
    if n is None:
        return 0.0
    try:
        for i in range(n.Elements.Count):
            e = n.Elements.Item(i)
            if e.Name.upper() == cid.upper():
                return float(e.Value) if e.Value else 0.0
    except Exception:
        pass
    return 0.0


print()
print('=== 面板 ===')
for x in MSGS[-14:]:
    print('   |', x[:170])
print()
print('=== 塔 ===')
for b in ['T-201', 'T-401', 'T-402', 'T-403', 'T-404']:
    bb = r'\Data\Blocks\%s\Output' % b
    print('  %-6s COND=%-11s REB=%-11s RR=%-7s Ttop=%-9s Tbot=%-9s' % (
        b, g(bb + r'\COND_DUTY'), g(bb + r'\REB_DUTY'), g(bb + r'\MOLE_RR'),
        g(bb + r'\TOP_TEMP'), g(bb + r'\BOTTOM_TEMP')))
print()
print('=== 关键流股 ===')
for s in ['S-110', 'S-113', 'S-114', 'S-115', 'S-116', 'S-117', 'S-118',
          'S-119', 'S-120', 'S-121', 'S-122']:
    b = r'\Data\Streams\%s\Output' % s
    print('  --- %s MASS=%s T=%s' % (s, g(b + r'\MASSFLMX\MIXED'), g(b + r'\TEMP_OUT')))
    nn = doc.Tree.FindNode(b + r'\MASSFLOW3')
    if nn is not None:
        try:
            for i2 in range(nn.Elements.Count):
                e = nn.Elements.Item(i2)
                try:
                    if e.Value and abs(float(e.Value)) > 0.05:
                        print('        %-8s %9.2f' % (e.Name, float(e.Value)))
                except Exception:
                    pass
        except Exception:
            pass
for p in [OUT.replace('.bkp', '_s.bkp')]:
    try:
        doc.SaveAs(p)
        print('已保存:', p, os.path.getsize(p))
    except Exception as ex:
        print('保存失败:', ex)
try:
    doc.Close()
except Exception:
    pass
print('DONE')
