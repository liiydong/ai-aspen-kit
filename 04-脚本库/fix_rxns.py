# -*- coding: utf-8 -*-
"""重修三个反应器 + T-401 进料板 + dump T-301 字段"""
import time, re
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\subs1.bkp'
OUT = r'D:\<化工工作区>\_probe\fix2.bkp'
NL = '\n'
SEP = ' \\ '


def rstoic(bid, temp, pres, stoic, convex, prod, icon='ICON1'):
    """stoic: [(reacno, cid, coef)]  convex: [(reacno, key, conv)]"""
    L = ['? BLOCK RSTOIC "%s" ? ; "METCBAR_MOLE" ; ; %s ; ' % (bid, icon)]
    L.append('\\ PARAM TEMP = %s <22> <4> PRES = %s <20> <5> SPEC-OPT = TP ' % (temp, pres))
    groups = []
    for rn, cid, co in stoic:
        groups.append('REACNO = %d STOIC-CID = "%s" STOIC-SSID = MIXED COEF = %s <0> <0>'
                      % (rn, cid, co))
    L.append('\\ \\ STOIC ' + ' / '.join(groups) + ' ')
    cg = []
    for rn, key, cv in convex:
        cg.append('EXT-REACNO = %d KEY-SSID = MIXED KEY-CID = "%s" CONV = %s <0> <0>'
                  % (rn, key, cv))
    L.append('\\ \\ CONVEX ' + ' / '.join(cg) + ' ')
    L.append('\\ \\ PRODUCTS SID = "%s" \\ ' % prod)
    return NL.join(L)


R101 = rstoic('R-101', '405.0', '1.8',
              [(1, '3-MP', -1.0), (1, 'NH3', -1.0), (1, 'O2', -1.5),
               (1, '3-CP', 1.0), (1, 'H2O', 3.0),
               (2, '3-MP', -1.0), (2, 'O2', -7.5), (2, 'CO2', 6.0),
               (2, 'H2O', 3.0), (2, 'HCN', 1.0)],
              [(1, '3-MP', '.86'), (2, '3-MP', '.02')], 'S-106')

R501 = rstoic('R-501', '140.0', '4.0',
              [(1, '3-CP', -1.0), (1, 'H2O', -1.0), (1, 'NAM', 1.0),
               (2, 'NAM', -1.0), (2, 'H2O', -1.0), (2, 'NAC', 1.0), (2, 'NH3', 1.0)],
              [(1, '3-CP', '.985'), (2, 'NAM', '.005')], 'S-201')

F501 = rstoic('F-501', '60.0', '1.0',
              [(1, 'NAOH', -2.0), (1, 'H2SO4', -1.0), (1, 'NA2SO4', 1.0), (1, 'H2O', 2.0)],
              [(1, 'H2SO4', '1.0')], 'S-301')

t = open(SRC, encoding='utf-8', errors='ignore').read()

# 替换三个 RSTOIC 段
starts = [m.start() for m in
          re.finditer(r'\?\s*BLOCK\s+[A-Z0-9]+\s+"?[A-Za-z0-9-]+"?\s*\?', t)]
NEW = {'R-101': R101, 'R-501': R501, 'F-501': F501}
rep = []
for idx in range(len(starts) - 1, -1, -1):
    s0 = starts[idx]
    s1 = starts[idx + 1] if idx + 1 < len(starts) else len(t)
    m = re.match(r'\?\s*BLOCK\s+RSTOIC\s+"([^"]+)"\s*\?', t[s0:s1])
    if m and m.group(1) in NEW:
        t = t[:s0] + NEW[m.group(1)] + NL + t[s1:]
        rep.append(m.group(1))
print('已替换 RSTOIC 段:', rep, flush=True)

# T-401 的两个进料都给板号
m = re.search(r'(BLOCK\s+BLKID\s*=\s*"T-401".*?)FEEDS\s+([^\n]*?)(\s*\\ \\ PRODUCTS)', t, re.S)
if m:
    old = m.group(2)
    new = ' FEED-SID = "S-114" FEED-STAGE = 12 / FEED-SID = "S-111" FEED-STAGE = 12 '
    t = t[:m.start(2)] + new + t[m.end(2):]
    print('T-401 FEEDS 改为:', re.sub(r'\s+', ' ', new), flush=True)
else:
    # 换一种方式：在 FEEDS 行里补 S-111 的板号
    t2 = re.sub(r'(FEED-SID = "S-114" FEED-STAGE = 12)(\s*\n\s*\\\s*\\\s*FEEDS\s*)',
                r'\1 / FEED-SID = "S-111" FEED-STAGE = 12\2', t)
    if t2 != t:
        t = t2
        print('T-401 FEEDS 已补 S-111 板号', flush=True)
    else:
        print('⚠ T-401 FEEDS 未改到，稍后用 COM 处理', flush=True)

open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)
print('已写出', OUT, flush=True)

# dump T-301 字段
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(OUT)
time.sleep(4)
n = doc.Tree.FindNode(r'\Data\Blocks\T-301\Input')
print()
print('=== T-301 (Extract) 输入字段中与 L1/L2/P-SPEC/T-EST 相关的 ===', flush=True)
for i in range(n.Elements.Count):
    ch = n.Elements.Item(i)
    u = ch.Name.upper()
    if any(k in u for k in ['L1', 'L2', 'P-SPEC', 'PSPEC', 'T-EST', 'TEST', 'PRES', 'TEMP', 'EST']):
        try:
            c = ch.Elements.Count
        except Exception:
            c = '-'
        try:
            v = repr(ch.Value)[:40]
        except Exception:
            v = ''
        print('   %-24s 子=%-4s %s' % (ch.Name, c, v), flush=True)
try:
    doc.Close()
except Exception:
    pass
