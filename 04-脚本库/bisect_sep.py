# -*- coding: utf-8 -*-
"""二分定位：到底哪一步改动导致 R-101 段丢失"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\_probe\bipA2.bkp'
base = open(SRC, encoding='utf-8', errors='ignore').read()

FS_OLD = ('BLKID = "T-301" BLKTYPE = "EXTRACT" MDLTYPE = "Extract" IN = ( "S-112" M0-1 '
          '\n"S-111" M1-2 "S-115" M2-3 ) OUT = ( "S-113" M1-2 "S-114" M0-1 )')
FS_NEW = ('BLKID = "T-301" BLKTYPE = "SEP" MDLTYPE = "Sep" IN = ( "S-112" M0-1 '
          '"S-111" M0-1 "S-115" M0-1 ) OUT = ( "S-113" M0-1 "S-114" M0-1 )')


def t301_sec(t, with_desc=True, with_param1=True):
    i = t.find('? COMPONENTS MAIN ?')
    j = t.find('? COMPONENTS "COMP-LIST"', i)
    cids = re.findall(r'CID = "?([A-Za-z0-9-]+)"?', t[i:j])
    FR = {'H2O': 0.999, 'TOL': 0.0005, '3-CP': 0.005, '4-CP': 0.005,
          '3-MP': 0.005, '4-MP': 0.005, 'HCN': 0.50, 'NH3': 0.99}
    recs = ['PARAM-STREAM = "S-113" SUBSTREAM = MIXED COMPS = "%s" FRACS = %s <0> <0> ' % (c, FR[c])
            for c in cids if c in FR]
    para = 'PARAM ' + ' /  '.join(recs) + ' \\ '
    head = '? BLOCK SEP "T-301" ? ; "METCBAR_MOLE" ; ; ICON2 ; \\ '
    if with_desc:
        head += 'DESCRIPTION DESCRIPTION = "Toluene extraction" \\ \\ '
    if with_param1:
        head += 'PARAM1 PRES1 = 1.0 <20> <5> \\ \\ '
    return head + para


def repl_sec(t, newsec):
    starts = [m.start() for m in re.finditer(r'\?\s*BLOCK\s+', t)]
    for k in range(len(starts) - 1, -1, -1):
        s0 = starts[k]
        s1 = starts[k + 1] if k + 1 < len(starts) else len(t)
        m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
        if m and m.group(2) == 'T-301':
            return t[:s0] + newsec + t[s1:]
    return t


def run(tag, path):
    doc = win32.DispatchEx('Apwn.Document')
    try:
        doc.SuppressDialogs = True
    except Exception:
        pass
    doc.InitFromArchive2(path)
    time.sleep(3)
    try:
        doc.Engine.Run2(False)
    except Exception as ex:
        print('%-34s Run2 异常 %s' % (tag, ex), flush=True)
    t0 = time.time()
    while time.time() - t0 < 180:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.25)
        if time.time() - t0 > 8:
            try:
                if not bool(doc.Engine.IsRunning):
                    break
            except Exception:
                break
    for _ in range(40):
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

    tot = g(r'\Data\Streams\S-121\Output\MASSFLMX\MIXED')
    print('%-34s  S-121 质量=%s' % (tag, tot), flush=True)
    try:
        doc.Close()
    except Exception:
        pass


# V1：只改 FLOWSHEET 类型
t = base.replace(FS_OLD, FS_NEW, 1)
open(r'D:\<化工工作区>\_probe\bi1.bkp', 'w', encoding='utf-8', errors='ignore').write(t)
run('V1 仅 FLOWSHEET 改类型', r'D:\<化工工作区>\_probe\bi1.bkp')

# V2：FLOWSHEET + 段落替换（保留原段落结构，仅换类型名与参数）
t = base.replace(FS_OLD, FS_NEW, 1)
t = repl_sec(t, t301_sec(base))
open(r'D:\<化工工作区>\_probe\bi2.bkp', 'w', encoding='utf-8', errors='ignore').write(t)
run('V2 + 新 SEP 段(带DESC/PARAM1)', r'D:\<化工工作区>\_probe\bi2.bkp')

# V3：FLOWSHEET + 最简 SEP 段（无 DESC / 无 PARAM1）
t = base.replace(FS_OLD, FS_NEW, 1)
t = repl_sec(t, t301_sec(base, with_desc=False, with_param1=False))
open(r'D:\<化工工作区>\_probe\bi3.bkp', 'w', encoding='utf-8', errors='ignore').write(t)
run('V3 + 最简 SEP 段', r'D:\<化工工作区>\_probe\bi3.bkp')

# V4：只改塔参数
t = base
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
            t = t[:s0] + seg2 + t[s1:]
            break
open(r'D:\<化工工作区>\_probe\bi4.bkp', 'w', encoding='utf-8', errors='ignore').write(t)
run('V4 仅塔参数(0.24bar/D:F0.9)', r'D:\<化工工作区>\_probe\bi4.bkp')

print('DONE')
