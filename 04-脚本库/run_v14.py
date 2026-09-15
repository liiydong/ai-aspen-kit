# -*- coding: utf-8 -*-
"""v14：萃取塔改 Sep（99.5%）+ 塔压按瑞邦实测重设 + 保留 BDBANK 修复"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\_probe\bipA2.bkp'      # 已含 BDBANK 修复
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_v14.bkp'

t = open(SRC, encoding='utf-8', errors='ignore').read()

# ---------- 1. 取组分清单 ----------
i = t.find('? COMPONENTS MAIN ?')
j = t.find('? COMPONENTS "COMP-LIST"', i)
cids = re.findall(r'CID = "?([A-Za-z0-9-]+)"?', t[i:j])
print('组分(%d):' % len(cids), cids)

# ---------- 2. 生成 SEP 的 FRACS（S-113 = 萃余水相） ----------
FRAC_TO_AQ = {
    'H2O': 0.999, 'TOL': 0.0005,
    '3-CP': 0.005, '4-CP': 0.005, '3-MP': 0.005, '4-MP': 0.005,
    'HCN': 0.50, 'NH3': 0.99,
    'NAM': 0.999, 'NAC': 0.999,
    'O2': 0.999, 'N2': 0.999, 'CO2': 0.999, 'CO': 0.999, 'AIR': 0.999,
    'NAOH': 0.999, 'H2SO4': 0.999, 'NA2SO4': 0.999,
}
recs = []
for c in cids:
    f = FRAC_TO_AQ.get(c, 0.999)
    recs.append('PARAM-STREAM = "S-113" SUBSTREAM = MIXED COMPS = "%s" FRACS = %s <0> <0> ' % (c, f))
sep_para = 'PARAM ' + ' /  '.join(recs) + ' \\ '

# ---------- 3. 替换 T-301 段 ----------
starts = [m.start() for m in re.finditer(r'\?\s*BLOCK\s+', t)]
t301 = None
for k in range(len(starts) - 1, -1, -1):
    s0 = starts[k]
    s1 = starts[k + 1] if k + 1 < len(starts) else len(t)
    m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
    if m and m.group(2) == 'T-301':
        t301 = (s0, s1)
        break
print('T-301 段位置:', t301)

new301 = ('? BLOCK SEP "T-301" ? ; "METCBAR_MOLE" ; ; ICON2 ; \\ '
          'DESCRIPTION DESCRIPTION = "甲苯逆流萃取（按瑞邦实测萃取率 99.5%）" \\ \\ '
          'PARAM1 PRES1 = 1.0 <20> <5> \\ \\ ' + sep_para + '\\ ')
t = t[:t301[0]] + new301 + t[t301[1]:]

# ---------- 4. FLOWSHEET 里改块类型 ----------
t2, n = re.subn(r'(BLOCK BLKID = "T-301" BLKTYPE = )"EXTRACT"( MDLTYPE = )"Extract"',
                r'\1"SEP"\2"Sep"', t)
print('FLOWSHEET 改块类型:', n)

# ---------- 5. 塔压按瑞邦实测重设 ----------
#   T-401 甲苯塔 0.24 bar / T-402 0.26 / T-403 0.4 / T-404 0.3
PRESS = {'T-401': 0.24, 'T-402': 0.26, 'T-403': 0.40, 'T-404': 0.30}
for bid, p in PRESS.items():
    ii = t2.find('BLOCK RADFRAC "%s"' % bid)
    if ii < 0:
        ii = t2.find('BLOCK RADFRAC \n"%s"' % bid)
    if ii < 0:
        print('  !! 未找到', bid)
        continue
    jj = t2.find('? BLOCK', ii + 10)
    seg = t2[ii:jj]
    seg2 = re.sub(r'PRES1 = [\d.]+ <20> <5>', 'PRES1 = %s <20> <5>' % p, seg, count=1)
    if seg2 != seg:
        t2 = t2[:ii] + seg2 + t2[jj:]
        print('  %s PRES1 -> %s bar' % (bid, p))
    else:
        print('  %s 未替换（原段无 PRES1 匹配）' % bid)

# ---------- 6. T-401 采出比：几乎全部甲苯走塔顶 ----------
t2, n = re.subn(r'("COL-SPECS" D:F = )[\d.]+( <-1> <0>)', r'\g<1>0.9300\g<2>', t2)
print('T-401 D:F 替换:', n)

open(OUT, 'w', encoding='utf-8', errors='ignore').write(t2)
print('写出:', OUT, os.path.getsize(OUT))

# ---------- 7. 运行 ----------
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
for x in MSGS[-18:]:
    print('   |', x[:170])

print()
print('=== 全流股 ===')
for s in ['S-104', 'S-106', 'S-109', 'S-110', 'S-112', 'S-113', 'S-114', 'S-115',
          'S-116', 'S-117', 'S-118', 'S-119', 'S-120', 'S-121', 'S-122']:
    b = r'\Data\Streams\%s\Output' % s
    print('  %-7s MASS=%-12s T=%-8s' % (s, g(b + r'\MASSFLMX\MIXED'), g(b + r'\TEMP_OUT')))

print()
print('=== 塔 ===')
for b in ['T-201', 'T-401', 'T-402', 'T-403', 'T-404']:
    bb = r'\Data\Blocks\%s\Output' % b
    print('  %-6s COND=%-12s REB=%-12s RR=%-8s Ttop=%-8s Tbot=%-8s' % (
        b, g(bb + r'\COND_DUTY'), g(bb + r'\REB_DUTY'), g(bb + r'\MOLE_RR'),
        g(bb + r'\TOP_TEMP'), g(bb + r'\BOTTOM_TEMP')))

print()
print('=== 关键流股组成 ===')
for s in ['S-113', 'S-114', 'S-116', 'S-117', 'S-121']:
    b = r'\Data\Streams\%s\Output' % s
    print('  --- %s MASS=%s' % (s, g(b + r'\MASSFLMX\MIXED')))
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
