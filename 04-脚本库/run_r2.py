# -*- coding: utf-8 -*-
"""R2：扩展「反应与吸收段」——修正插入锚点"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_q3_s.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_r2.bkp'
t = open(SRC, encoding='utf-8', errors='ignore').read()
print('源大小', len(t))

# 锚点确认
A_FS = t.find('"DEF-STREAM"')
A_STR = t.find('? SETUP MAIN ?')
A_BLOCK_FIRST = t.find('? BLOCK')
A_DSET = t.find('\nDSET')
A_PROPS = t.find('? PROPERTIES MAIN ?')
A_DSET = A_BLOCK_FIRST   # 无 DSET 标记，改为插在首个 BLOCK 段之前
print('锚点: DEF-STREAM=%s SETUP_MAIN=%s 首个BLOCK=%s PROPERTIES_MAIN=%s' % (
    A_FS, A_STR, A_BLOCK_FIRST, A_PROPS))
assert A_FS > 0 and A_STR > 0 and A_BLOCK_FIRST > 0


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


# ---------- 1) 注册表 ----------
NEW = [('C-101', 'Compr', 'COMPR'), ('E-102', 'Heater', 'HEATER'), ('E-103', 'Heater', 'HEATER'),
       ('E-105', 'Heater', 'HEATER'), ('T-202', 'RadFrac', 'RADFRAC')]
m = re.search(r';\s*\n(\d+)\s*\n>VERSION 0', t)
n_old = int(m.group(1))
reg_add = ''.join('>VERSION 0\n%s\n%s\nBuilt-In\n%s\n' % e for e in NEW)
t = t[:A_STR] + reg_add + t[A_STR:]
t = t[:m.start()] + '; \n%d\n>VERSION 0' % (n_old + len(NEW)) + t[m.end():]
print('1) 注册表 %d -> %d ✓' % (n_old, n_old + len(NEW)))

# ---------- 2) FLOWSHEET ----------
i, j = slice_entry(t, 'M-101')
seg = t[i:j]
seg2 = re.sub(r'IN = \([^)]*\)', 'IN = ( "S-101" M0-1 "S-126" M0-1 "S-128" M0-1 )', seg, count=1)
t = t[:i] + seg2 + t[j:]
print('2a) M-101 ->', repr(seg2[:130]))
i, j = slice_entry(t, 'E-104')
seg = t[i:j]
seg2 = seg.replace('"S-106"', '"S-129"')
t = t[:i] + seg2 + t[j:]
print('2b) E-104 ->', repr(seg2[:130]))

A_FS = t.find('"DEF-STREAM"')     # 重新定位（前面的编辑已改变偏移）
def ent(bid, bt, mt, ins_, outs_):
    return ('BLOCK BLKID = "%s" BLKTYPE = "%s" MDLTYPE = "%s" IN = ( %s ) OUT = ( %s ) \\ \\ '
            % (bid, bt, mt, ins_, outs_))
ins = (ent('C-101', 'COMPR', 'Compr', '"S-103" M0-1', '"S-127" M0-1')
       + ent('E-102', 'HEATER', 'Heater', '"S-127" M0-1', '"S-128" M0-1')
       + ent('E-103', 'HEATER', 'Heater', '"S-102" M0-1', '"S-126" M0-1')
       + ent('E-105', 'HEATER', 'Heater', '"S-106" M0-1', '"S-129" M0-1')
       + ent('T-202', 'RADFRAC', 'RadFrac', '"S-109" M0-1 "S-132" M0-1', '"S-130" M1-2 "S-131" M2-3'))
t = t[:A_FS] + ins + t[A_FS:]
print('2c) 追加 5 个 FLOWSHEET 条目 ✓')

# ---------- 3) 新流股 S-132（插在 BLOCK 段之前） ----------
A_BLOCK_FIRST = t.find('? BLOCK')
s132 = wr('? STREAM MATERIAL "S-132" ? ; "METCBAR_MOLE" ; \\ SUBSTREAM SSID = MIXED TEMP = 25 <22> <4> '
          'PRES = 1.0 <20> <5> BASIS = "MOLE-FLOW" MIXED-SPEC = TP TOTAL = 20.0 <-89> <0> JUNK = 1 \\ \\ '
          'MOLE-FLOW SSID1 = MIXED CID = H2O FLOW = 20.0 <-89> <3> \\ ')
t = t[:A_BLOCK_FIRST] + s132 + '\n' + t[A_BLOCK_FIRST:]
print('3) 新增 S-132 洗涤水 ✓')

# ---------- 4) 新块段落（插在首个已有 BLOCK 段之前） ----------
A_DSET = t.find('? BLOCK')
SECT = []
SECT.append(wr('? BLOCK COMPR "C-101" ? ; "METCBAR_MOLE" ; ; ICON2 ; \\ '
               'DESCRIPTION DESCRIPTION = "Air blower" \\ \\ '
               'PARAM TYPE = ISENTROPIC OPT-SPEC = PRES PRES = 2.0 <20> <5> SEFF = 0.72 <0> <0> \\ '))
for bid, temp, desc in [('E-102', '150.0', 'Air preheater'),
                        ('E-103', '60.0', 'Ammonia vaporizer'),
                        ('E-105', '300.0', 'Waste heat boiler')]:
    pres = '2.0' if bid != 'E-105' else '1.8'
    SECT.append(wr('? BLOCK HEATER "%s" ? ; "METCBAR_MOLE" ; ; HEATER ; \\ '
                   'DESCRIPTION DESCRIPTION = "%s" \\ \\ '
                   'PARAM TEMP = %s <22> <4> PRES = %s <20> <5> SPEC-OPT = TP \\ ' % (bid, desc, temp, pres)))
T202 = ('? BLOCK RADFRAC "T-202" ? ; "METCBAR_MOLE" ; ; ABSBR1 ; \\ PARAM NSTAGE = 6 NSTAGEMAX = 7 \\ \\ '
        'PARAM2 \\ \\ "COL-CONFIG" CONDENSER = NONE REBOILER = KETTLE \\ \\ '
        'FEEDS FEED-SID = "S-132" FEED-STAGE = 1 /  FEED-SID = "S-109" FEED-STAGE = 6 \\ \\ '
        'PRODUCTS PROD-STREAM = "S-130" PROD-STAGE = 1 PROD-PHASE = V P-S = N /  '
        'PROD-STREAM = "S-131" PROD-STAGE = 6 PROD-PHASE = L P-S = N \\ \\ '
        '"P-SPEC2" PRES1 = 1.0 <20> <5> \\ \\ "COL-SPECS" BASIS-RDV = 1.0 <0> <0> BASIS-BR = 0.1 <-1> <0> \\ \\ '
        'T-EST TEMP-STAGE = 1 TEMP-EST = 40.0 <22> <4> /  TEMP-STAGE = 6 TEMP-EST = 60.0 <22> <4> \\ \\ '
        '"KLL-VECS" \\ \\ "TRSZ-VECS" \\ \\ "PCKSR-VECS" \\ ')
SECT.append(wr(T202))
t = t[:A_DSET + 1] + '\n'.join(SECT) + t[A_DSET + 1:]
print('4) 新增 5 个块段落 ✓')

assert len(re.findall(r'\?\s*BLOCK\s+RADFRAC', t)) == 5, re.findall(r'\?\s*BLOCK\s+RADFRAC', t)
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
    if re.search(r'ERROR|SEVERE|Summary|Terminal|completed|Warnings|Warning', x):
        print('   |', x[:165])
print()
print('=== 新块 ===')
for b in ['C-101', 'E-102', 'E-103', 'E-105', 'T-202']:
    bb = r'\Data\Blocks\%s\Output' % b
    print('  %-6s T=%-9s P=%-9s DUTY=%-11s COND=%-9s REB=%s' % (
        b, g(bb + r'\TEMP'), g(bb + r'\PRES'), g(bb + r'\DUTY'), g(bb + r'\COND_DUTY'), g(bb + r'\REB_DUTY')))
print()
print('=== 流股 ===')
for s in ['S-103', 'S-127', 'S-102', 'S-126', 'S-104', 'S-106', 'S-129', 'S-108',
          'S-109', 'S-130', 'S-131', 'S-110', 'S-121']:
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
