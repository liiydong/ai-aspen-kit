# -*- coding: utf-8 -*-
"""验证：把 SEP 段落按 76 字符换行后是否正常"""
import sys, re, time, textwrap
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

base = open(r'D:\<化工工作区>\_probe\bipA2.bkp', encoding='utf-8', errors='ignore').read()
FS_OLD = ('BLKID = "T-301" BLKTYPE = "EXTRACT" MDLTYPE = "Extract" IN = ( "S-112" M0-1 '
          '\n"S-111" M1-2 "S-115" M2-3 ) OUT = ( "S-113" M1-2 "S-114" M0-1 )')
REG_OLD = 'T-301\nExtract\nBuilt-In\nEXTRACT\n'


def wr(s, w=74):
    """按 74 字符换行，尽量在空格处断"""
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
            return t[:s0] + wr(para) + t[s1:]
    return t


MSGS = []


class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            MSGS.append(s)


i = base.find('? COMPONENTS MAIN ?')
j = base.find('? COMPONENTS "COMP-LIST"', i)
cids = re.findall(r'CID = "?([A-Za-z0-9-]+)"?', base[i:j])
FR = {'H2O': 0.999, 'TOL': 0.0005, '3-CP': 0.005, '4-CP': 0.005,
      '3-MP': 0.005, '4-MP': 0.005, 'HCN': 0.50, 'NH3': 0.99}
recs = ['PARAM-STREAM = "S-113" SUBSTREAM = MIXED COMPS = "%s" FRACS = %s <0> <0> ' % (c, FR[c])
        for c in cids if c in FR]
para = ('? BLOCK SEP "T-301" ? ; "METCBAR_MOLE" ; ; ICON1 ; \\ '
        'DESCRIPTION DESCRIPTION = "Toluene extraction 99.5pct" \\ \\ '
        'PARAM1 PRES1 = 1.0 <20> <5> \\ \\ '
        'PARAM ' + ' /  '.join(recs) + ' \\ ')
p = r'D:\<化工工作区>\_probe\sc_wrap.bkp'
open(p, 'w', encoding='utf-8', errors='ignore').write(build(para))
print('段落换行后:')
print(wr(para))
print()

doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
win32.WithEvents(doc, Sink)
doc.InitFromArchive2(p)
time.sleep(3)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 200:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 8:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break


def g(pp):
    n = doc.Tree.FindNode(pp)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return None


sev = [x for x in MSGS if 'SEVERE' in x]
print('严重报错:', len(sev))
for x in sev[:4]:
    print('  |', x[:150])
print()
print('=== 流股 ===')
for s in ['S-113', 'S-114', 'S-121']:
    b = r'\Data\Streams\%s\Output' % s
    print('  %s MASS=%s' % (s, g(b + r'\MASSFLMX\MIXED')))
    nn = doc.Tree.FindNode(b + r'\MASSFLOW3')
    if nn is not None:
        try:
            for i2 in range(nn.Elements.Count):
                e = nn.Elements.Item(i2)
                try:
                    if e.Value and abs(float(e.Value)) > 0.05:
                        print('      %-8s %9.2f' % (e.Name, float(e.Value)))
                except Exception:
                    pass
        except Exception:
            pass
try:
    doc.Close()
except Exception:
    pass
print('DONE')
