# -*- coding: utf-8 -*-
"""v10：剔除 NaOH/H2SO4/Na2SO4（占 0.03%）+ 删除 F-501/S-123/S-125 + 物性方法改回 NRTL"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_v7.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_v10.bkp'

t = open(SRC, encoding='utf-8', errors='ignore').read()
REM = ['NAOH', 'H2SO4', 'NA2SO4']

print('=== 1) 删 MOLEC-STRUCT 段 ===', flush=True)
for cid in REM:
    pat = re.compile(r'\?\s*PROPERTIES\s+"MOLEC-STRUCT"\s+"?%s"?\s*\?' % re.escape(cid))
    k = len(pat.findall(t))
    t = pat.sub('', t)
    print('   %-8s 删除 %d 处' % (cid, k), flush=True)

print('=== 2) 从 COMPONENTS MAIN 删除组分条目 ===', flush=True)
m = re.search(r'\?\s*COMPONENTS MAIN\s*\?', t)
mb = re.search(r'\n\?\s*[A-Z]', t[m.end():])
body_end = m.end() + mb.start()
body = t[m.end():body_end]
ents = re.split(r'\s*/\s*', body)
kept = []
for e in ents:
    cid = None
    mm = re.search(r'CID\s*=\s*"?([A-Za-z0-9_-]+)"?', e)
    if mm:
        cid = mm.group(1).strip('"')
    if cid in REM:
        print('   删除组分:', cid, flush=True)
        continue
    kept.append(e)
t = t[:m.end()] + ' / '.join(kept) + t[body_end:]

print('=== 3) 删 FLOWSHEET 里的 F-501 记录 / 改 R-501、E-601 进料 ===', flush=True)
patF = re.compile(r'\\\s*BLOCK\s+BLKID\s*=\s*"F-501"[\s\S]{0,260}?OUT\s*=\s*\([^)]*\)')
print('   F-501 记录数:', len(patF.findall(t)), flush=True)
t = patF.sub('', t)
t2 = re.sub(r'(BLKID\s*=\s*"R-501"[\s\S]{0,200}?IN\s*=\s*\(\s*[^)]*?)\s*"S-123"\s*M0-1', r'\1', t)
print('   R-501 去 S-123:', t2 != t, flush=True)
t = t2
t2 = re.sub(r'(BLKID\s*=\s*"E-601"[\s\S]{0,200}?IN\s*=\s*\(\s*)"S-301"', r'\1"S-201"', t)
print('   E-601 改 S-201:', t2 != t, flush=True)
t = t2

print('=== 4) 删 F-501 的 BLOCK 段 ===', flush=True)
for bid in ['F-501']:
    m = re.search(r'\?\s*BLOCK\s+RSTOIC\s+"?%s\s*"?\s*\?' % bid, t)
    if m:
        nx = re.search(r'\n\?\s*BLOCK\s', t[m.start() + 10:])
        end = m.start() + 10 + nx.start() if nx else len(t)
        t = t[:m.start()] + t[end + 1:]
        print('   删除 BLOCK 段:', bid, flush=True)

print('=== 5) 删 S-123 / S-125 / S-301 的 STREAM 段 ===', flush=True)
for sid in ['S-123', 'S-125', 'S-301']:
    pat = re.compile(r'\?\s*STREAM\s+MATERIAL\s+"?%s"?\s*\?' % re.escape(sid))
    m = pat.search(t)
    if not m:
        print('   未找到', sid, flush=True)
        continue
    nx = re.search(r'\n\?\s*', t[m.end():])
    end = m.end() + nx.start() if nx else len(t)
    t = t[:m.start()] + t[end + 1:]
    print('   删除 STREAM 段:', sid, flush=True)

print('=== 6) 物性方法改回 NRTL ===', flush=True)
m3 = re.search(r'PARAM\s+BASE\s*=\s*"[^"]*"', t)
t = t[:m3.start()] + 'PARAM BASE = "NRTL"' + t[m3.end():]

open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)
print('写出:', OUT, len(t), flush=True)

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
try:
    win32.WithEvents(doc, Sink)
except Exception:
    pass
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

print()
print('=== 面板 ===', flush=True)
for i, x in enumerate(MSGS[:70], 1):
    print('%3d| %s' % (i, x[:175]), flush=True)


def gg(p):
    n = doc.Tree.FindNode(p)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return None


print()
print('=== 关键流股：3-CP 去了哪 ===', flush=True)
for s in ['S-106', 'S-109', 'S-110', 'S-112', 'S-113', 'S-114', 'S-116', 'S-118', 'S-120', 'S-121']:
    b = r'\Data\Streams\%s\Output' % s
    print('  --- %s  MASS=%s  VFRAC=%s' % (
        s, gg(b + r'\MASSFLMX\MIXED'), gg(b + r'\VFRAC_OUT\MIXED')), flush=True)
    n = doc.Tree.FindNode(b + r'\MASSFLOW3')
    if n is not None:
        for i in range(n.Elements.Count):
            try:
                e = n.Elements.Item(i)
                if e.Value and abs(float(e.Value)) > 0.05:
                    print('        %-8s %9.2f' % (e.Name, float(e.Value)), flush=True)
            except Exception:
                pass

print()
print('=== 塔结果 ===', flush=True)
for b in ['T-201', 'T-301', 'T-401', 'T-402', 'T-403', 'T-404']:
    bb = r'\Data\Blocks\%s\Output' % b
    print('  %-7s COND=%s REB=%s RR=%s Ttop=%s Tbot=%s' % (
        b, gg(bb + r'\COND_DUTY'), gg(bb + r'\REB_DUTY'), gg(bb + r'\MOLE_RR'),
        gg(bb + r'\TOP_TEMP'), gg(bb + r'\BOTTOM_TEMP')), flush=True)

try:
    doc.SaveAs(OUT.replace('.bkp', '.apwz'))
except Exception:
    pass
try:
    doc.Close()
except Exception:
    pass
print('DONE', flush=True)
