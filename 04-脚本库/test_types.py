# -*- coding: utf-8 -*-
"""对照测试：换块类型本身会不会触发 R-101 报错"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

base = open(r'D:\<化工工作区>\_probe\bipA2.bkp', encoding='utf-8', errors='ignore').read()

FS_OLD = ('BLKID = "T-301" BLKTYPE = "EXTRACT" MDLTYPE = "Extract" IN = ( "S-112" M0-1 '
          '\n"S-111" M1-2 "S-115" M2-3 ) OUT = ( "S-113" M1-2 "S-114" M0-1 )')
REG_OLD = 'T-301\nExtract\nBuilt-In\nEXTRACT\n'

SEC_OLD_MARK = '? BLOCK EXTRACT "T-301"'


def build(model, blktype, para):
    t = base
    t = t.replace(REG_OLD, 'T-301\n%s\nBuilt-In\n%s\n' % (model, blktype), 1)
    fs = ('BLKID = "T-301" BLKTYPE = "%s" MDLTYPE = "%s" IN = ( "S-112" M0-1 '
          '"S-111" M0-1 "S-115" M0-1 ) OUT = ( "S-113" M0-1 "S-114" M0-1 )' % (blktype, model))
    t = t.replace(FS_OLD, fs, 1)
    # 替换段
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

    bad = [x for x in MSGS if 'R-101' in x or 'NO BLOCK PARAGRAPH' in x or 'SEVERE' in x]
    print('%-28s  S-121=%s   报错:%d  %s' % (tag, g(r'\Data\Streams\S-121\Output\MASSFLMX\MIXED'),
                                            len(bad), (bad[0][:80] if bad else '')), flush=True)
    try:
        doc.Close()
    except Exception:
        pass


cases = [
    ('EXTRACT 原样', 'Extract', 'EXTRACT',
     '? BLOCK EXTRACT "T-301" ? ; "METCBAR_MOLE" ; ; ICON2 ; \\ PARAM NSTAGE = 8 \\ '),
    ('HEATER', 'Heater', 'HEATER',
     '? BLOCK HEATER "T-301" ? ; "METCBAR_MOLE" ; ; ICON1 ; \\ PARAM TEMP = 60.0 <22> <4> PRES = 1.0 <20> <5> \\ '),
    ('SEP', 'Sep', 'SEP',
     '? BLOCK SEP "T-301" ? ; "METCBAR_MOLE" ; ; ICON1 ; \\ PARAM PARAM-STREAM = "S-113" SUBSTREAM = MIXED COMPS = "H2O" FRACS = 0.999 <0> <0> \\ '),
    ('FSPLIT', 'FSplit', 'FSPLIT',
     '? BLOCK FSPLIT "T-301" ? ; "METCBAR_MOLE" ; ; ICON1 ; \\ FRAC 0.5 \\ '),
]
for tag, model, blktype, para in cases:
    p = r'D:\<化工工作区>\_probe\tc_%s.bkp' % blktype.lower()
    open(p, 'w', encoding='utf-8', errors='ignore').write(build(model, blktype, para))
    run(tag, p)
print('DONE')
