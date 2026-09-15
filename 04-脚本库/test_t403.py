# -*- coding: utf-8 -*-
"""决定性测试：T-403 改回 RadFrac(45板块) 看 3-CP/4-CP 分离"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_最终定稿.bkp'
OUT = r'D:\<化工工作区>\_probe\t403_radfrac.bkp'
t = open(SRC, encoding='utf-8', errors='ignore').read()
L = []

# 1) 文件头注册表
t2, n1 = re.subn(r'\nT-403\nSep\nBuilt-In\nSEP\n>VERSION 0\n',
                 '\nT-403\nRadFrac\nBuilt-In\nRADFRAC\n>VERSION 0\n', t)
L.append('1) 注册表替换 %d 处' % n1)

# 2) FLOWSHEET
t2, n2 = re.subn(
    r'BLOCK\s+BLKID\s*=\s*"T-403"\s+BLKTYPE\s*=\s*"SEP"\s+MDLTYPE\s*=\s*"Sep"\s+'
    r'IN\s*=\s*\(\s*"S-118"\s+M0-1\s*\)\s+OUT\s*=\s*\(\s*"S-119"\s+M0-1\s+"S-120"\s+M0-1\s*\)',
    '"BLOCK"', t2)
L.append('2) FLOWSHEET 替换 %d 处' % n2)
# 用函数式替换避免转义问题
def fs_rep(m):
    return ('BLOCK BLKID = "T-403" BLKTYPE = "RADFRAC" MDLTYPE = "RadFrac" '
            '\\ \\ IN = ( "S-118" M0-1 ) OUT = ( "S-119" M1-2 "S-120" M2-3 )')
# 先还原刚才的错误替换
t2 = t2.replace('BLKID = "T-403" BLKTYPE = "SEP" MDLTYPE = "Sep" \\ \\ IN = ( "S-118" M0-1 ) OUT = ( "S-119" M0-1 "S-120" M0-1 )', '<<<T403>>>')
# 稳妥：用正则定位并整体替换
m = re.search(r'BLOCK\s+BLKID\s*=\s*"T-403".{0,200}?"S-120"\s+M0-1\s*\)', t2, re.S)
if m:
    t2 = t2[:m.start()] + fs_rep(None) + t2[m.end():]
    L.append('2b) FLOWSHEET 记录已重写 (%d 字符 -> %d)' % (m.end() - m.start(), len(fs_rep(None))))
else:
    L.append('2b) FLOWSHEET 未定位')

# 3) 段头与段落
m = re.search(r'\?\s*BLOCK\s+SEP\s+"T-403"\s*\?', t2)
if m:
    nxt = re.search(r'\n\?\s*BLOCK\s', t2[m.end():])
    e = m.end() + nxt.start() if nxt else m.end() + 3000
    new_sec = (
        '? BLOCK\nRADFRAC "T-403" ? ; "METCBAR_MOLE" ; ; FRACT1 ; \\\n'
        'PARAM NSTAGE = 45 NSTAGEMAX = 46 \\ \\\n'
        'PARAM2 \\ \\\n'
        '"COL-CONFIG" CONDENSER = TOTAL REBOILER = KETTLE \\ \\\n'
        'FEEDS FEED-SID = "S-118" FEED-STAGE = 22 \\ \\\n'
        'PRODUCTS PROD-STREAM = "S-119"\nPROD-STAGE = 1 PROD-PHASE = L P-S = N /\n'
        'PROD-STREAM = "S-120"\nPROD-STAGE = 45 PROD-PHASE = L P-S = N \\ \\\n'
        '"P-SPEC2" PRES1 = 0.40\n<20> <5> \\ \\\n'
        '"COL-SPECS" D:F = 0.0050\n<-1> <0> BASIS-RDV = 0.0 <0> <0> BASIS-RR = 18.0\n<-1> <0> D:F-BASIS = MOLE \\ \\\n'
        'T-EST TEMP-STAGE = 1 TEMP-EST = 165.0\n<22> <4> /\nTEMP-STAGE = 45 TEMP-EST = 172.0\n<22> <4> \\ \\\n'
        '"KLL-VECS" \\ \\\n"TRSZ-VECS" \\ \\\n"PCKSR-VECS" \\\n'
    )
    t2 = t2[:m.start()] + new_sec + t2[e:]
    L.append('3) 段落已替换为 RadFrac(45 板块)')
else:
    L.append('3) 段落未定位')

# 4) DSET
t2, n4 = re.subn(r'BLOCK SEP T-403', 'BLOCK RADFRAC T-403', t2)
L.append('4) DSET 替换 %d 处' % n4)

open(OUT, 'w', encoding='utf-8', errors='ignore').write(t2)
L.append('写出 %s  %d 字节' % (OUT, os.path.getsize(OUT)))

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
while time.time() - t0 < 600:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 10:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(80):
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

def comp(s):
    nn = doc.Tree.FindNode(r'\Data\Streams\%s\Output\MASSFLOW3' % s)
    d = {}
    if nn is None:
        return d
    try:
        for i in range(nn.Elements.Count):
            e = nn.Elements.Item(i)
            try:
                v = float(e.Value) if e.Value is not None else 0.0
                if abs(v) > 1e-6:
                    d[e.Name] = v
            except Exception:
                pass
    except Exception:
        pass
    return d

L.append('')
L.append('===== 控制面板关键行 =====')
for i, x in enumerate(MSGS):
    if re.search(r'Terminal|Severe|Errors|Warnings|completed|T-403|ERROR|error', x):
        L.append('  %d | %s' % (i, x[:170]))

L.append('')
L.append('===== T-403 (RadFrac) 结果 =====')
for k in ['TOP_TEMP', 'BOTTOM_TEMP', 'RR', 'D:F', 'COND_DUTY', 'REB_DUTY', 'TOP_LFLOW']:
    L.append('  %-12s = %s' % (k, g(r'\Data\Blocks\T-403\Output\%s' % k)))
for s in ['S-118', 'S-119', 'S-120']:
    c = comp(s)
    tt = sum(c.values())
    L.append('  %s = %.2f kg/h : %s' % (s, tt, ', '.join('%s %.4f' % (k, v) for k, v in sorted(c.items(), key=lambda x: -x[1])[:6])))
    if tt:
        m4 = c.get('4-CP', 0); m3 = c.get('3-CP', 0)
        L.append('        3-CP %.2f kg/h (%.3f%%)  4-CP %.2f kg/h (%.3f%%)' % (m3, m3 / tt * 100, m4, m4 / tt * 100))
open(r'D:\<化工工作区>\_probe\t403_test.txt', 'w', encoding='utf-8').write('\n'.join(L))
try:
    doc.SaveAs(r'D:\<化工工作区>\_probe\t403_radfrac_s.bkp')
except Exception as ex:
    L.append('save fail %s' % ex)
try:
    doc.Close()
except Exception:
    pass
print('DONE')
