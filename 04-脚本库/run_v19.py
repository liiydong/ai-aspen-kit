# -*- coding: utf-8 -*-
"""v19 最终版：SEP 萃取 + 合理甲苯补充量 + 瑞邦实测塔压"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\_probe\bipA2.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_v19.bkp'
t = open(SRC, encoding='utf-8', errors='ignore').read()


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


# 1) 注册表
t = t.replace('T-301\nExtract\nBuilt-In\nEXTRACT\n', 'T-301\nSep\nBuilt-In\nSEP\n', 1)
# 2) FLOWSHEET
FS_OLD = ('BLKID = "T-301" BLKTYPE = "EXTRACT" MDLTYPE = "Extract" IN = ( "S-112" M0-1 '
          '\n"S-111" M1-2 "S-115" M2-3 ) OUT = ( "S-113" M1-2 "S-114" M0-1 )')
FS_NEW = ('BLKID = "T-301" BLKTYPE = "SEP" MDLTYPE = "Sep" IN = ( "S-112" M0-1 '
          '"S-111" M0-1 "S-115" M0-1 ) OUT = ( "S-113" M0-1 "S-114" M0-1 )')
assert FS_OLD in t
t = t.replace(FS_OLD, FS_NEW, 1)

# 3) SEP 段（带换行）
i = t.find('? COMPONENTS MAIN ?')
j = t.find('? COMPONENTS "COMP-LIST"', i)
cids = re.findall(r'CID = "?([A-Za-z0-9-]+)"?', t[i:j])
FR = {'H2O': 0.999, 'TOL': 0.0005, '3-CP': 0.005, '4-CP': 0.005,
      '3-MP': 0.005, '4-MP': 0.005, 'HCN': 0.50, 'NH3': 0.99}
recs = ['PARAM-STREAM = "S-113" SUBSTREAM = MIXED COMPS = "%s" FRACS = %s <0> <0> ' % (c, FR[c])
        for c in cids if c in FR]
para = ('? BLOCK SEP "T-301" ? ; "METCBAR_MOLE" ; ; ICON2 ; \\ '
        'DESCRIPTION DESCRIPTION = "Toluene counter-current extraction, 99.5 pct per Ruibang" \\ \\ '
        'PARAM1 PRES1 = 1.0 <20> <5> \\ \\ PARAM ' + ' /  '.join(recs) + ' \\ ')
starts = [m.start() for m in re.finditer(r'\?\s*BLOCK\s+', t)]
loc = None
for k in range(len(starts) - 1, -1, -1):
    s0 = starts[k]
    s1 = starts[k + 1] if k + 1 < len(starts) else len(t)
    m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
    if m and m.group(2) == 'T-301':
        loc = (s0, s1)
        break
t = t[:loc[0]] + wr(para) + t[loc[1]:]

# 4) S-111 补充甲苯 -> 0.25 kmol/h
i = t.find('?  STREAM MATERIAL "S-111"')
if i < 0:
    i = t.find('? STREAM MATERIAL "S-111"')
j2 = t.find('? STREAM', i + 20)
seg = t[i:j2]
seg2 = re.sub(r'TOTAL = 24\.47', 'TOTAL = 0.25', seg)
seg2 = re.sub(r'(CID = TOL FLOW = )24\.47', r'\g<1>0.25', seg2)
if seg2 == seg:
    seg2 = seg.replace('24.47', '0.25')
print('S-111 改动:', 'OK' if seg2 != seg else '未改')
t = t[:i] + seg2 + t[j2:]

# 5) 塔参数
for bid, pres, df in [('T-401', '0.24', '0.9950'), ('T-402', '0.26', None),
                      ('T-403', '0.40', None), ('T-404', '0.30', None)]:
    starts = [m.start() for m in re.finditer(r'\?\s*BLOCK\s+', t)]
    for k in range(len(starts) - 1, -1, -1):
        s0 = starts[k]
        s1 = starts[k + 1] if k + 1 < len(starts) else len(t)
        m = re.match(r'\?\s*BLOCK\s+RADFRAC\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
        if m and m.group(1) == bid:
            sg = t[s0:s1]
            sg2 = re.sub(r'PRES1 = [\d.]+ <20> <5>', 'PRES1 = %s <20> <5>' % pres, sg, count=1)
            if df:
                sg2 = re.sub(r'D:F = [\d.]+ <-1> <0>', 'D:F = %s <-1> <0>' % df, sg2, count=1)
            if sg2 != sg:
                t = t[:s0] + sg2 + t[s1:]
                print('  %s: PRES1=%s%s' % (bid, pres, ' D:F=' + df if df else ''))
            break

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
try:
    win32.WithEvents(doc, Sink)
except Exception:
    pass
doc.InitFromArchive2(OUT)
time.sleep(3)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 420:
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
print('=== 面板尾部 ===')
for x in MSGS[-20:]:
    print('   |', x[:170])
print()
print('=== 塔 ===')
for b in ['T-201', 'T-401', 'T-402', 'T-403', 'T-404']:
    bb = r'\Data\Blocks\%s\Output' % b
    print('  %-6s COND=%-11s REB=%-11s RR=%-7s Ttop=%-8s Tbot=%-8s' % (
        b, g(bb + r'\COND_DUTY'), g(bb + r'\REB_DUTY'), g(bb + r'\MOLE_RR'),
        g(bb + r'\TOP_TEMP'), g(bb + r'\BOTTOM_TEMP')))
print()
print('=== 流股 ===')
for s in ['S-104', 'S-106', 'S-109', 'S-110', 'S-112', 'S-113', 'S-114', 'S-115',
          'S-116', 'S-117', 'S-118', 'S-119', 'S-120', 'S-121', 'S-122', 'S-301']:
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
for p in [OUT.replace('.bkp', '_s.bkp'), OUT.replace('.bkp', '.apwz')]:
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
