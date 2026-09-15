# -*- coding: utf-8 -*-
"""细分：SEP 段落里哪种写法触发 R-101 报错"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

base = open(r'D:\<化工工作区>\_probe\bipA2.bkp', encoding='utf-8', errors='ignore').read()
FS_OLD = ('BLKID = "T-301" BLKTYPE = "EXTRACT" MDLTYPE = "Extract" IN = ( "S-112" M0-1 '
          '\n"S-111" M1-2 "S-115" M2-3 ) OUT = ( "S-113" M1-2 "S-114" M0-1 )')
REG_OLD = 'T-301\nExtract\nBuilt-In\nEXTRACT\n'


def build(para):
    t = base
    t = t.replace(REG_OLD, 'T-301\nSep\nBuilt-In\nSEP\n', 1)
    t = t.replace(FS_OLD, ('BLKID = "T-301" BLKTYPE = "SEP" MDLTYPE = "Sep" IN = ( "S-112" M0-1 '
                           '"S-111" M0-1 "S-115" M0-1 ) OUT = ( "S-113" M0-1 "S-114" M0-1 )'), 1)
    starts = [m.start() for m in re.finditer(r'\?\s*BLOCK\s+', t)]
    for k in range(len(starts) - 1, -1, -1):
        s0 = starts[k]
        s1 = starts[k + 1] if k + 1 < len(starts) else len(t)
        m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
        if m and m.group(2) == 'T-301':
            return t[:s0] + para + t[s1:]
    return t


def run(tag, path):
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
    win32.WithEvents(doc, Sink)
    doc.InitFromArchive2(path)
    time.sleep(3)
    try:
        doc.Engine.Run2(False)
    except Exception:
        pass
    t0 = time.time()
    while time.time() - t0 < 150:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.25)
        if time.time() - t0 > 8:
            try:
                if not bool(doc.Engine.IsRunning):
                    break
            except Exception:
                break

    def g(p):
        n = doc.Tree.FindNode(p)
        if n is None:
            return None
        try:
            return n.Value
        except Exception:
            return None

    sev = [x for x in MSGS if 'SEVERE' in x or 'NO BLOCK PARAGRAPH' in x]
    print('%-34s S-121=%-14s 严重报错:%d' % (tag, g(r'\Data\Streams\S-121\Output\MASSFLMX\MIXED'), len(sev)), flush=True)
    for x in sev[:3]:
        print('       |', x[:150], flush=True)
    try:
        doc.Close()
    except Exception:
        pass


H = '? BLOCK SEP "T-301" ? ; "METCBAR_MOLE" ; ; ICON1 ; \\ '
comp_all = ('PARAM PARAM-STREAM = "S-113" SUBSTREAM = MIXED COMPS = "3-MP" FRACS = 0.005 <0> <0>  /  '
            'PARAM-STREAM = "S-113" SUBSTREAM = MIXED COMPS = "NH3" FRACS = 0.99 <0> <0>  /  '
            'PARAM-STREAM = "S-113" SUBSTREAM = MIXED COMPS = "3-CP" FRACS = 0.005 <0> <0>  /  '
            'PARAM-STREAM = "S-113" SUBSTREAM = MIXED COMPS = "4-CP" FRACS = 0.005 <0> <0>  /  '
            'PARAM-STREAM = "S-113" SUBSTREAM = MIXED COMPS = "H2O" FRACS = 0.999 <0> <0>  /  '
            'PARAM-STREAM = "S-113" SUBSTREAM = MIXED COMPS = "TOL" FRACS = 0.0005 <0> <0>  /  '
            'PARAM-STREAM = "S-113" SUBSTREAM = MIXED COMPS = "HCN" FRACS = 0.5 <0> <0> \\ ')
comp_1 = 'PARAM PARAM-STREAM = "S-113" SUBSTREAM = MIXED COMPS = "H2O" FRACS = 0.999 <0> <0> \\ '

cases = [
    ('A 单组分', H + comp_1),
    ('B 全组分', H + comp_all),
    ('C 单组分+DESC', H + 'DESCRIPTION DESCRIPTION = "Tol ext" \\ \\ ' + comp_1),
    ('D 单组分+PARAM1', H + 'PARAM1 PRES1 = 1.0 <20> <5> \\ \\ ' + comp_1),
    ('E 全组分+DESC', H + 'DESCRIPTION DESCRIPTION = "Tol ext" \\ \\ ' + comp_all),
    ('F 全组分+PARAM1', H + 'PARAM1 PRES1 = 1.0 <20> <5> \\ \\ ' + comp_all),
]
for tag, para in cases:
    p = r'D:\<化工工作区>\_probe\sc_%s.bkp' % tag.split()[0].lower()
    open(p, 'w', encoding='utf-8', errors='ignore').write(build(para))
    run(tag, p)
print('DONE')
