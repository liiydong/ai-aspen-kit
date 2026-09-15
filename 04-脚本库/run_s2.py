# -*- coding: utf-8 -*-
"""S2：扩展后处理段（修正流股名冲突 / 进料板 / FEEDS 名）"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_r3_s.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_s2.bkp'
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


NEW = [('T-302', 'RadFrac', 'RADFRAC'), ('M-501', 'Mixer', 'MIXER'), ('M-502', 'Sep', 'SEP'),
       ('E-602', 'Flash2', 'FLASH2'), ('E-603', 'Flash2', 'FLASH2'), ('C-601', 'Flash2', 'FLASH2'),
       ('M-601', 'Sep', 'SEP'), ('D-601', 'Flash2', 'FLASH2'), ('P-301', 'Pump', 'PUMP'),
       ('P-401', 'Pump', 'PUMP')]

# 1) 注册表
m = re.search(r';\s*\n(\d+)\s*\n>VERSION 0', t)
n_old = int(m.group(1))
A_STR = t.find('? SETUP MAIN ?')
t = t[:A_STR] + ''.join('>VERSION 0\n%s\n%s\nBuilt-In\n%s\n' % e for e in NEW) + t[A_STR:]
t = t[:m.start()] + '; \n%d\n>VERSION 0' % (n_old + len(NEW)) + t[m.end():]
print('1) 注册表 %d -> %d ✓' % (n_old, n_old + len(NEW)))

# 2) 重接
t = set_in(t, 'E-601', '"S-203" M0-1')
t = set_in(t, 'T-401', '"S-114A" M0-1')
t = set_in(t, 'T-402', '"S-116A" M0-1')
# FEEDS 里的流股名同步
t = repl_block(t, 'T-401', lambda s: s.replace('FEED-SID = "S-114"', 'FEED-SID = "S-114A"'))
t = repl_block(t, 'T-402', lambda s: s.replace('FEED-SID = "S-116"', 'FEED-SID = "S-116A"'))
print('2) E-601/T-401/T-402 重接 + FEEDS 同步 ✓')

# 3) FLOWSHEET 追加
A_FS = t.find('"DEF-STREAM"')
def ent(bid, bt, mt, ins_, outs_):
    return ('BLOCK BLKID = "%s" BLKTYPE = "%s" MDLTYPE = "%s" IN = ( %s ) OUT = ( %s ) \\ \\ '
            % (bid, bt, mt, ins_, outs_))
ins = (ent('T-302', 'RADFRAC', 'RadFrac', '"S-113" M0-1', '"S-212" M1-2 "S-213" M2-3')
       + ent('M-501', 'MIXER', 'Mixer', '"S-201" M0-1 "S-210" M0-1', '"S-202" M0-1')
       + ent('M-502', 'SEP', 'Sep', '"S-202" M0-1', '"S-203" M0-1 "S-211" M0-1')
       + ent('E-602', 'FLASH2', 'Flash2', '"S-402" M0-1', '"S-214" M1-2 "S-205" M0-1')
       + ent('E-603', 'FLASH2', 'Flash2', '"S-205" M0-1', '"S-403" M1-2 "S-206" M0-1')
       + ent('C-601', 'FLASH2', 'Flash2', '"S-206" M0-1', '"S-404" M1-2 "S-207" M0-1')
       + ent('M-601', 'SEP', 'Sep', '"S-207" M0-1', '"S-405" M0-1 "S-208" M0-1')
       + ent('D-601', 'FLASH2', 'Flash2', '"S-208" M0-1', '"S-406" M1-2 "S-209" M0-1')
       + ent('P-301', 'PUMP', 'Pump', '"S-114" M0-1', '"S-114A" M0-1')
       + ent('P-401', 'PUMP', 'Pump', '"S-116" M0-1', '"S-116A" M0-1'))
t = t[:A_FS] + wr(ins) + t[A_FS:]
print('3) FLOWSHEET 追加 10 条 ✓')

# 4) 新进料 S-210
A_BLOCK = t.find('? BLOCK')
s210 = wr('? STREAM MATERIAL "S-210" ? ; "METCBAR_MOLE" ; \\ SUBSTREAM SSID = MIXED TEMP = 80 <22> <4> '
          'PRES = 1.0 <20> <5> BASIS = "MOLE-FLOW" MIXED-SPEC = TP TOTAL = 5.0 <-89> <0> JUNK = 1 \\ \\ '
          'MOLE-FLOW SSID1 = MIXED CID = H2O FLOW = 5.0 <-89> <3> \\ ')
t = t[:A_BLOCK] + s210 + '\n' + t[A_BLOCK:]
print('4) S-210 ✓')

# 5) 新块段落
ci = t.find('? COMPONENTS MAIN ?')
cj = t.find('? COMPONENTS "COMP-LIST"', ci)
cids = re.findall(r'CID = "?([A-Za-z0-9-]+)"?', t[ci:cj])


def sep_para(bid, out_s, fr):
    recs = ['PARAM-STREAM = "%s" SUBSTREAM = MIXED COMPS = "%s" FRACS = %s <0> <0> ' % (out_s, c, fr.get(c, 0.002))
            for c in cids]
    return ('? BLOCK SEP "%s" ? ; "METCBAR_MOLE" ; ; ICON1 ; \\ PARAM1 PRES1 = 1.0 <20> <5> \\ \\ '
            'PARAM ' % bid) + ' /  '.join(recs) + ' \\ '


def flash(bid, temp, pres, desc):
    return ('? BLOCK FLASH2 "%s" ? ; "METCBAR_MOLE" ; ; "V-DRUM1" ; \\ '
            'DESCRIPTION DESCRIPTION = "%s" \\ \\ '
            'PARAM TEMP = %s <22> <4> PRES = %s <20> <5> \\ \\ FRAC SUBSTREAM = MIXED \\ '
            % (bid, desc, temp, pres))


T302 = ('? BLOCK RADFRAC "T-302" ? ; "METCBAR_MOLE" ; ; ABSBR1 ; \\ PARAM NSTAGE = 8 NSTAGEMAX = 9 \\ \\ '
        'PARAM2 \\ \\ "COL-CONFIG" CONDENSER = NONE REBOILER = KETTLE \\ \\ '
        'FEEDS FEED-SID = "S-113" FEED-STAGE = 1 \\ \\ '
        'PRODUCTS PROD-STREAM = "S-212" PROD-STAGE = 1 PROD-PHASE = V P-S = N /  '
        'PROD-STREAM = "S-213" PROD-STAGE = 8 PROD-PHASE = L P-S = N \\ \\ '
        '"P-SPEC2" PRES1 = 1.0 <20> <5> \\ \\ "COL-SPECS" BASIS-RDV = 1.0 <0> <0> BASIS-BR = 0.1 <-1> <0> \\ \\ '
        'T-EST TEMP-STAGE = 1 TEMP-EST = 40.0 <22> <4> /  TEMP-STAGE = 8 TEMP-EST = 100.0 <22> <4> \\ \\ '
        '"KLL-VECS" \\ \\ "TRSZ-VECS" \\ \\ "PCKSR-VECS" \\ ')

S = [wr(T302)]
S.append(wr('? BLOCK MIXER "M-501" ? ; "METCBAR_MOLE" ; ; TRIANGLE ; \\ '
            'DESCRIPTION DESCRIPTION = "Decolorizing vessel" \\ \\ PARAM PRES = 1. <20> <5> VISITED = 1 \\ '))
S.append(wr(sep_para('M-502', 'S-211', {})))
S.append(wr(flash('E-602', '60.', '0.20', 'Second effect evaporator')))
S.append(wr(flash('E-603', '50.', '0.10', 'Third effect evaporator')))
S.append(wr(flash('C-601', '30.', '0.05', 'Crystallizer')))
S.append(wr(sep_para('M-601', 'S-405', {'H2O': 0.90, 'NAM': 0.05, 'NAC': 0.05, '3-CP': 0.95, '3-MP': 0.95})))
S.append(wr(flash('D-601', '110.', '0.10', 'Dryer')))
S.append(wr('? BLOCK PUMP "P-301" ? ; "METCBAR_MOLE" ; ; ICON1 ; \\ '
            'DESCRIPTION DESCRIPTION = "Extract pump" \\ \\ '
            'PARAM PRES = 1.5 <20> <5> EFF = 0.7 <0> <0> OPT-SPEC = PRES \\ '))
S.append(wr('? BLOCK PUMP "P-401" ? ; "METCBAR_MOLE" ; ; ICON1 ; \\ '
            'DESCRIPTION DESCRIPTION = "Crude CP pump" \\ \\ '
            'PARAM PRES = 1.0 <20> <5> EFF = 0.7 <0> <0> OPT-SPEC = PRES \\ '))
A_BLOCK = t.find('? BLOCK')
t = t[:A_BLOCK] + '\n'.join(S) + t[A_BLOCK:]
print('5) 10 个新块段落 ✓')

assert len(re.findall(r'\?\s*BLOCK\s+RADFRAC', t)) == 6
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
print('=== 全部面板消息 ===')
for i, x in enumerate(MSGS[:80], 1):
    print('%3d| %s' % (i, x[:165]))
print()
print('=== 后处理段流股 ===')
for s in ['S-113', 'S-201', 'S-202', 'S-203', 'S-204', 'S-205', 'S-206', 'S-207', 'S-208',
          'S-209', 'S-211', 'S-212', 'S-213', 'S-401', 'S-402', 'S-403', 'S-404', 'S-405',
          'S-406', 'S-214', 'S-114', 'S-114A', 'S-116', 'S-116A', 'S-121']:
    print('  %-7s MASS=%-11s' % (s, g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s)))
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
