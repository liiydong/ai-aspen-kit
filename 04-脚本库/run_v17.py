# -*- coding: utf-8 -*-
"""v17：修正终止符（多一个反斜杠吞掉 R-101 段的问题）"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\_probe\bipA2.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_v17.bkp'
t = open(SRC, encoding='utf-8', errors='ignore').read()

# 1) FLOWSHEET
old_fs = ('BLKID = "T-301" BLKTYPE = "EXTRACT" MDLTYPE = "Extract" IN = ( "S-112" M0-1 '
          '\n"S-111" M1-2 "S-115" M2-3 ) OUT = ( "S-113" M1-2 "S-114" M0-1 )')
new_fs = ('BLKID = "T-301" BLKTYPE = "SEP" MDLTYPE = "Sep" IN = ( "S-112" M0-1 '
          '"S-111" M0-1 "S-115" M0-1 ) OUT = ( "S-113" M0-1 "S-114" M0-1 )')
assert old_fs in t, 'FLOWSHEET 串未匹配'
t = t.replace(old_fs, new_fs, 1)
print('FLOWSHEET OK')

# 2) SEP 参数段（结尾只保留一个 \ ）
i = t.find('? COMPONENTS MAIN ?')
j = t.find('? COMPONENTS "COMP-LIST"', i)
cids = re.findall(r'CID = "?([A-Za-z0-9-]+)"?', t[i:j])
FR = {'H2O': 0.999, 'TOL': 0.0005, '3-CP': 0.005, '4-CP': 0.005,
      '3-MP': 0.005, '4-MP': 0.005, 'HCN': 0.50, 'NH3': 0.99}
recs = ['PARAM-STREAM = "S-113" SUBSTREAM = MIXED COMPS = "%s" FRACS = %s <0> <0> ' % (c, FR[c])
        for c in cids if c in FR]
sep_para = 'PARAM ' + ' /  '.join(recs) + ' \\ '

starts = [m.start() for m in re.finditer(r'\?\s*BLOCK\s+', t)]
loc = None
for k in range(len(starts) - 1, -1, -1):
    s0 = starts[k]
    s1 = starts[k + 1] if k + 1 < len(starts) else len(t)
    m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
    if m and m.group(2) == 'T-301':
        loc = (s0, s1)
        break
new301 = ('? BLOCK SEP "T-301" ? ; "METCBAR_MOLE" ; ; ICON2 ; \\ '
          'DESCRIPTION DESCRIPTION = "Toluene extraction 99.5pct per Ruibang" \\ \\ '
          'PARAM1 PRES1 = 1.0 <20> <5> \\ \\ ' + sep_para)   # 不再追加多余反斜杠
t = t[:loc[0]] + new301 + t[loc[1]:]
print('T-301 段已替换，长', len(new301))
print('  尾部:', repr(new301[-90:]))
print('  下一段:', repr(t[loc[0] + len(new301):loc[0] + len(new301) + 60]))

# 3) 塔参数
for bid, pres, df in [('T-401', '0.24', '0.9000'), ('T-402', '0.26', None),
                      ('T-403', '0.40', None), ('T-404', '0.30', None)]:
    starts = [m.start() for m in re.finditer(r'\?\s*BLOCK\s+', t)]
    for k in range(len(starts) - 1, -1, -1):
        s0 = starts[k]
        s1 = starts[k + 1] if k + 1 < len(starts) else len(t)
        m = re.match(r'\?\s*BLOCK\s+RADFRAC\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
        if m and m.group(1) == bid:
            seg = t[s0:s1]
            seg2 = re.sub(r'PRES1 = [\d.]+ <20> <5>', 'PRES1 = %s <20> <5>' % pres, seg, count=1)
            if df:
                seg2 = re.sub(r'D:F = [\d.]+ <-1> <0>', 'D:F = %s <-1> <0>' % df, seg2, count=1)
            if seg2 != seg:
                t = t[:s0] + seg2 + t[s1:]
                print('  %s PRES1=%s%s' % (bid, pres, ' D:F=' + df if df else ''))
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
try:
    doc.SuppressDialogs = True
except Exception:
    pass
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
for x in MSGS[-22:]:
    print('   |', x[:170])

print()
print('=== 塔 ===')
for b in ['T-201', 'T-401', 'T-402', 'T-403', 'T-404']:
    bb = r'\Data\Blocks\%s\Output' % b
    print('  %-6s COND=%-11s REB=%-11s RR=%-7s Ttop=%-8s Tbot=%-8s' % (
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

try:
    doc.SaveAs(OUT.replace('.bkp', '_s.bkp'))
    print('已保存:', OUT.replace('.bkp', '_s.bkp'))
except Exception as ex:
    print('保存失败:', ex)
try:
    doc.Close()
except Exception:
    pass
print('DONE')
