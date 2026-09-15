# -*- coding: utf-8 -*-
"""R3：扩展「反应与吸收段」（修正 HEATER/COMPR 段落写法）"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_q3_s.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_r3.bkp'
t = open(SRC, encoding='utf-8', errors='ignore').read()
print('源大小', len(t))


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
    assert i > 0, blkid
    j = x.find('BLOCK', i + 20)
    return i, j


# 1) 注册表
NEW = [('C-101', 'Compr', 'COMPR'), ('E-102', 'Heater', 'HEATER'), ('E-103', 'Heater', 'HEATER'),
       ('E-105', 'Heater', 'HEATER'), ('T-202', 'RadFrac', 'RADFRAC')]
m = re.search(r';\s*\n(\d+)\s*\n>VERSION 0', t)
n_old = int(m.group(1))
A_STR = t.find('? SETUP MAIN ?')
reg_add = ''.join('>VERSION 0\n%s\n%s\nBuilt-In\n%s\n' % e for e in NEW)
t = t[:A_STR] + reg_add + t[A_STR:]
t = t[:m.start()] + '; \n%d\n>VERSION 0' % (n_old + len(NEW)) + t[m.end():]
print('1) 注册表 %d -> %d ✓' % (n_old, n_old + len(NEW)))

# 2) FLOWSHEET
i, j = slice_entry(t, 'M-101')
seg = t[i:j]
seg2 = re.sub(r'IN = \([^)]*\)', 'IN = ( "S-101" M0-1 "S-126" M0-1 "S-128" M0-1 )', seg, count=1)
t = t[:i] + seg2 + t[j:]
i, j = slice_entry(t, 'E-104')
seg = t[i:j]
seg2 = seg.replace('"S-106"', '"S-129"')
t = t[:i] + seg2 + t[j:]
print('2a) M-101 / E-104 重接 ✓')

A_FS = t.find('"DEF-STREAM"')
def ent(bid, bt, mt, ins_, outs_):
    return ('BLOCK BLKID = "%s" BLKTYPE = "%s" MDLTYPE = "%s" IN = ( %s ) OUT = ( %s ) \\ \\ '
            % (bid, bt, mt, ins_, outs_))
ins = (ent('C-101', 'COMPR', 'Compr', '"S-103" M0-1', '"S-127" M0-1')
       + ent('E-102', 'HEATER', 'Heater', '"S-127" M0-1', '"S-128" M0-1')
       + ent('E-103', 'HEATER', 'Heater', '"S-102" M0-1', '"S-126" M0-1')
       + ent('E-105', 'HEATER', 'Heater', '"S-106" M0-1', '"S-129" M0-1')
       + ent('T-202', 'RADFRAC', 'RadFrac', '"S-109" M0-1 "S-132" M0-1', '"S-130" M1-2 "S-131" M2-3'))
t = t[:A_FS] + wr(ins) + t[A_FS:]
print('2b) FLOWSHEET 追加 5 条（已按 74 字符换行）✓')

# 3) 空气进料改为常压（让空压机有用）
mi = t.find('? STREAM MATERIAL "S-103"')
if mi < 0:
    mi = re.search(r'\?\s*STREAM\s+MATERIAL\s*\n?\s*"S-103"', t).start()
mj = t.find('?  STREAM', mi + 20)
if mj < 0:
    mj = t.find('? STREAM', mi + 20)
seg = t[mi:mj]
seg2 = re.sub(r'PRES = [\d.]+ <20> <5>', 'PRES = 1.0 <20> <5>', seg, count=1)
t = t[:mi] + seg2 + t[mj:]
print('3) S-103 空气改为 1.0 bar ✓')

# 4) S-132 洗涤水（插在首个 BLOCK 段前）
A_BLOCK = t.find('? BLOCK')
s132 = wr('? STREAM MATERIAL "S-132" ? ; "METCBAR_MOLE" ; \\ SUBSTREAM SSID = MIXED TEMP = 25 <22> <4> '
          'PRES = 1.0 <20> <5> BASIS = "MOLE-FLOW" MIXED-SPEC = TP TOTAL = 20.0 <-89> <0> JUNK = 1 \\ \\ '
          'MOLE-FLOW SSID1 = MIXED CID = H2O FLOW = 20.0 <-89> <3> \\ ')
t = t[:A_BLOCK] + s132 + '\n' + t[A_BLOCK:]
print('4) S-132 洗涤水 ✓')

# 5) 新块段落
A_BLOCK = t.find('? BLOCK')
S = []
S.append(wr('? BLOCK COMPR "C-101" ? ; "METCBAR_MOLE" ; ; ICON2 ; \\ '
            'DESCRIPTION DESCRIPTION = "Air blower" \\ \\ '
            'PARAM TYPE = ISENTROPIC OPT-SPEC = DELP DELP = 1.0 <20> <5> SEFF = 0.72 <0> <0> \\ '))
S.append(wr('? BLOCK HEATER "E-102" ? ; "METCBAR_MOLE" ; ; HEATER ; \\ '
            'DESCRIPTION DESCRIPTION = "Air preheater" \\ \\ '
            'PARAM TEMP = 150. <22> <4> PRES = 2.0 <20> <5> \\ '))
S.append(wr('? BLOCK HEATER "E-103" ? ; "METCBAR_MOLE" ; ; HEATER ; \\ '
            'DESCRIPTION DESCRIPTION = "Ammonia vaporizer" \\ \\ '
            'PARAM TEMP = 60. <22> <4> PRES = 2.0 <20> <5> \\ '))
S.append(wr('? BLOCK HEATER "E-105" ? ; "METCBAR_MOLE" ; ; HEATER ; \\ '
            'DESCRIPTION DESCRIPTION = "Waste heat boiler" \\ \\ '
            'PARAM TEMP = 300. <22> <4> PRES = 1.8 <20> <5> \\ '))
S.append(wr('? BLOCK RADFRAC "T-202" ? ; "METCBAR_MOLE" ; ; ABSBR1 ; \\ PARAM NSTAGE = 6 NSTAGEMAX = 7 \\ \\ '
            'PARAM2 \\ \\ "COL-CONFIG" CONDENSER = NONE REBOILER = KETTLE \\ \\ '
            'FEEDS FEED-SID = "S-132" FEED-STAGE = 1 /  FEED-SID = "S-109" FEED-STAGE = 6 \\ \\ '
            'PRODUCTS PROD-STREAM = "S-130" PROD-STAGE = 1 PROD-PHASE = V P-S = N /  '
            'PROD-STREAM = "S-131" PROD-STAGE = 6 PROD-PHASE = L P-S = N \\ \\ '
            '"P-SPEC2" PRES1 = 1.0 <20> <5> \\ \\ "COL-SPECS" BASIS-RDV = 1.0 <0> <0> BASIS-BR = 0.1 <-1> <0> \\ \\ '
            'T-EST TEMP-STAGE = 1 TEMP-EST = 40.0 <22> <4> /  TEMP-STAGE = 6 TEMP-EST = 60.0 <22> <4> \\ \\ '
            '"KLL-VECS" \\ \\ "TRSZ-VECS" \\ \\ "PCKSR-VECS" \\ '))
t = t[:A_BLOCK] + '\n'.join(S) + t[A_BLOCK:]
print('5) 5 个新块段落 ✓')

assert len(re.findall(r'\?\s*BLOCK\s+RADFRAC', t)) == 5
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
print('=== 面板 ===')
for x in MSGS:
    if re.search(r'ERROR|SEVERE|Summary|Terminal|completed|Warning', x):
        print('   |', x[:165])
print()
print('=== 新块 ===')
for b in ['C-101', 'E-102', 'E-103', 'E-105', 'T-202']:
    bb = r'\Data\Blocks\%s\Output' % b
    print('  %-6s T=%-9s P=%-9s DUTY=%-11s COND=%-9s REB=%s' % (
        b, g(bb + r'\TEMP'), g(bb + r'\PRES'), g(bb + r'\DUTY'), g(bb + r'\COND_DUTY'), g(bb + r'\REB_DUTY')))
print()
print('=== 流股 ===')
for s in ['S-103', 'S-127', 'S-128', 'S-102', 'S-126', 'S-104', 'S-106', 'S-129',
          'S-108', 'S-109', 'S-130', 'S-131', 'S-110', 'S-121']:
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
