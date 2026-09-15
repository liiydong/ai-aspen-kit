# -*- coding: utf-8 -*-
"""R1：扩展「反应与吸收段」——加 C-101 空压机 / E-102 空气预热 / E-103 氨汽化 / E-105 废热锅炉 / T-202 尾气洗涤塔"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_q3_s.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_r1.bkp'
t = open(SRC, encoding='utf-8', errors='ignore').read()
print('源大小', len(t), '| RADFRAC段', len(re.findall(r'\?\s*BLOCK\s+RADFRAC', t)),
      '| 注册表块数', re.search(r';\s*\n(\d+)\s*\n>VERSION 0', t).group(1))


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


def slice_entry(x, blkid):
    i = x.find('BLKID = "%s"' % blkid)
    assert i > 0, '未找到 ' + blkid
    j = x.find('BLOCK', i + 20)
    if j < 0:
        j = x.find('\\\\', i + 20)
    return i, j


# ---------- 1) 注册表 ----------
NEW_BLOCKS = [
    ('C-101', 'Compr', 'COMPR', '空气鼓风机（加压至 2.0 bar）'),
    ('E-102', 'Heater', 'HEATER', '空气预热器（150 ℃）'),
    ('E-103', 'Heater', 'HEATER', '氨汽化器（60 ℃）'),
    ('E-105', 'Heater', 'HEATER', '反应气废热锅炉（冷却至 300 ℃，产蒸汽）'),
    ('T-202', 'RadFrac', 'RADFRAC', '尾气洗涤塔（二级吸收）'),
]
m = re.search(r';\s*\n(\d+)\s*\n>VERSION 0', t)
n_old = int(m.group(1))
n_new = n_old + len(NEW_BLOCKS)
reg_add = ''.join('>VERSION 0\n%s\n%s\nBuilt-In\n%s\n' % (a, b, c) for a, b, c, _ in NEW_BLOCKS)
k = t.find('? SETUP MAIN ?')
t = t[:k] + reg_add + t[k:]
t = t[:m.start()] + '; \n%d\n>VERSION 0' % n_new + t[m.end():]
print('1) 注册表 %d -> %d，新增 %d 个块 ✓' % (n_old, n_new, len(NEW_BLOCKS)))

# ---------- 2) FLOWSHEET 重接 ----------
# M-101 进料改为 S-101 / S-126(氨汽化后) / S-128(空气预热后)
i, j = slice_entry(t, 'M-101')
seg = t[i:j]
print('   M-101 原:', repr(seg))
seg2 = re.sub(r'IN = \([^)]*\)', 'IN = ( "S-101" M0-1 "S-126" M0-1 "S-128" M0-1 )', seg, count=1)
t = t[:i] + seg2 + t[j:]
print('   M-101 新:', repr(seg2))
# E-104 进料改为 S-129（废热锅炉出口）
i, j = slice_entry(t, 'E-104')
seg = t[i:j]
seg2 = seg.replace('"S-106"', '"S-129"')
t = t[:i] + seg2 + t[j:]
print('   E-104 新:', repr(seg2))
# 追加 5 个新块条目
last = [mm.start() for mm in re.finditer(r'BLOCK\s+BLKID', t)][-1]
e = t.find('\\\\ \\\\', last)
ins = ''
def ent(bid, bt, mt, ins_, outs_):
    return 'BLOCK BLKID = "%s" BLKTYPE = "%s" MDLTYPE = "%s" IN = ( %s ) OUT = ( %s ) \\ \\ ' % (bid, bt, mt, ins_, outs_)
ins += ent('C-101', 'COMPR', 'Compr', '"S-103" M0-1', '"S-127" M0-1')
ins += ent('E-102', 'HEATER', 'Heater', '"S-127" M0-1', '"S-128" M0-1')
ins += ent('E-103', 'HEATER', 'Heater', '"S-102" M0-1', '"S-126" M0-1')
ins += ent('E-105', 'HEATER', 'Heater', '"S-106" M0-1', '"S-129" M0-1')
ins += ent('T-202', 'RADFRAC', 'RadFrac', '"S-109" M0-1 "S-132" M0-1', '"S-130" M1-2 "S-131" M2-3')
t = t[:e + 5] + ins + t[e + 5:]
print('2) FLOWSHEET 追加 5 个块条目 ✓')

# ---------- 3) 新增块段落 ----------
SECT = []
SECT.append(wr('? BLOCK COMPR "C-101" ? ; "METCBAR_MOLE" ; ; ICON2 ; \\ '
               'DESCRIPTION DESCRIPTION = "Air blower" \\ \\ '
               'PARAM TYPE = ISENTROPIC OPT-SPEC = PRES PRES = 2.0 <20> <5> SEFF = 0.72 <0> <0> \\ '))
SECT.append(wr('? BLOCK HEATER "E-102" ? ; "METCBAR_MOLE" ; ; HEATER ; \\ '
               'DESCRIPTION DESCRIPTION = "Air preheater" \\ \\ '
               'PARAM TEMP = 150.0 <22> <4> PRES = 2.0 <20> <5> \\ '))
SECT.append(wr('? BLOCK HEATER "E-103" ? ; "METCBAR_MOLE" ; ; HEATER ; \\ '
               'DESCRIPTION DESCRIPTION = "Ammonia vaporizer" \\ \\ '
               'PARAM TEMP = 60.0 <22> <4> PRES = 2.0 <20> <5> \\ '))
SECT.append(wr('? BLOCK HEATER "E-105" ? ; "METCBAR_MOLE" ; ; HEATER ; \\ '
               'DESCRIPTION DESCRIPTION = "Reactor effluent waste heat boiler" \\ \\ '
               'PARAM TEMP = 300.0 <22> <4> PRES = 1.8 <20> <5> \\ '))
T202 = ('? BLOCK RADFRAC "T-202" ? ; "METCBAR_MOLE" ; ; ABSBR1 ; \\ PARAM NSTAGE = 6 NSTAGEMAX = 7 \\ \\ '
        'PARAM2 \\ \\ "COL-CONFIG" CONDENSER = NONE REBOILER = KETTLE \\ \\ '
        'FEEDS FEED-SID = "S-132" FEED-STAGE = 1 /  FEED-SID = "S-109" FEED-STAGE = 6 \\ \\ '
        'PRODUCTS PROD-STREAM = "S-130" PROD-STAGE = 1 PROD-PHASE = V P-S = N /  '
        'PROD-STREAM = "S-131" PROD-STAGE = 6 PROD-PHASE = L P-S = N \\ \\ '
        '"P-SPEC2" PRES1 = 1.0 <20> <5> \\ \\ "COL-SPECS" BASIS-RDV = 1.0 <0> <0> BASIS-BR = 0.1 <-1> <0> \\ \\ '
        'T-EST TEMP-STAGE = 1 TEMP-EST = 40.0 <22> <4> /  TEMP-STAGE = 6 TEMP-EST = 60.0 <22> <4> \\ \\ '
        '"KLL-VECS" \\ \\ "TRSZ-VECS" \\ \\ "PCKSR-VECS" \\ ')
SECT.append(wr(T202))
pk = t.find('? PROPERTIES MAIN ?')
t = t[:pk] + '\n'.join(SECT) + t[pk:]
print('3) 新增 5 个块段落 ✓')

# ---------- 4) 新增洗涤水进料流股 S-132 ----------
si = t.find('? STREAM')
stream132 = wr('? STREAM MATERIAL "S-132" ? ; "METCBAR_MOLE" ; \\ SUBSTREAM SSID = MIXED TEMP = 25 <22> <4> '
               'PRES = 1.0 <20> <5> BASIS = "MOLE-FLOW" MIXED-SPEC = TP TOTAL = 20.0 <-89> <0> JUNK = 1 \\ \\ '
               'MOLE-FLOW SSID1 = MIXED CID = H2O FLOW = 20.0 <-89> <3> \\ ')
t = t[:si] + stream132 + '\n' + t[si:]
print('4) 新增 S-132 洗涤水进料 ✓')

assert len(re.findall(r'\?\s*BLOCK\s+RADFRAC', t)) == 5
assert len(re.findall(r'BLOCK\s+RADFRAC\s*\n?\s*"T-401"', t)) == 1
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)
print('写出:', OUT, os.path.getsize(OUT))

# ---------- 5) 运行 ----------
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


def g(p):
    n = doc.Tree.FindNode(p)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return None


print()
print('=== 面板（含错误）===')
for x in MSGS:
    if re.search(r'ERROR|SEVERE|Summary|Terminal|completed|Warnings', x):
        print('   |', x[:165])
print()
print('=== 新块结果 ===')
for b in ['C-101', 'E-102', 'E-103', 'E-105', 'T-202']:
    bb = r'\Data\Blocks\%s\Output' % b
    print('  %-6s T=%-9s P=%-9s DUTY=%-11s COND=%-10s REB=%s' % (
        b, g(bb + r'\TEMP'), g(bb + r'\PRES'), g(bb + r'\DUTY'),
        g(bb + r'\COND_DUTY'), g(bb + r'\REB_DUTY')))
print()
print('=== 相关流股 ===')
for s in ['S-103', 'S-127', 'S-128', 'S-102', 'S-126', 'S-101', 'S-104',
          'S-106', 'S-129', 'S-108', 'S-109', 'S-130', 'S-131', 'S-110']:
    print('  %-7s MASS=%-11s T=%s' % (s, g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s),
                                      g(r'\Data\Streams\%s\Output\TEMP_OUT' % s)))
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
