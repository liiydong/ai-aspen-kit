# -*- coding: utf-8 -*-
"""v12：稳健删段 —— 用"下一个段头"定位，去掉 NaOH/H2SO4/Na2SO4 与 F-501/S-123，物性改 NRTL"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_v7.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_v12.bkp'
REM = {'NAOH', 'H2SO4', 'NA2SO4'}

t = open(SRC, encoding='utf-8', errors='ignore').read()
n0 = len(t)
print('源 %d 字符, ? BLOCK 段数=%d' % (n0, len(re.findall(r'\?\s*BLOCK\s', t))), flush=True)

# 1) MOLEC-STRUCT
for cid in REM:
    t = re.sub(r'\?\s*PROPERTIES\s+"MOLEC-STRUCT"\s+"?%s"?\s*\?' % re.escape(cid), '', t)

# 2) COMPONENTS MAIN
m = re.search(r'\?\s*COMPONENTS MAIN\s*\?', t)
mb = re.search(r'\?\s*[A-Z][A-Z0-9_/ -]{0,40}\s*\?', t[m.end():])
body_end = m.end() + mb.start()
body = t[m.end():body_end]
ents = re.split(r'\s*/\s*', body)
kept = []
for e in ents:
    mm = re.search(r'CID\s*=\s*"?([A-Za-z0-9_-]+)"?', e)
    cid = mm.group(1).strip('"') if mm else None
    if cid in REM:
        print('  删组分', cid, flush=True)
        continue
    kept.append(e)
t = t[:m.end()] + ' / '.join(kept) + t[body_end:]

# 3) FLOWSHEET
t = re.sub(r'\\\s*BLOCK\s+BLKID\s*=\s*"F-501"[\s\S]{0,300}?OUT\s*=\s*\([^)]*\)', '', t)
t = re.sub(r'(BLKID\s*=\s*"R-501"[\s\S]{0,240}?IN\s*=\s*\(\s*[^)]*?)\s*"S-123"\s*M0-1', r'\1', t)
t = re.sub(r'(BLKID\s*=\s*"E-601"[\s\S]{0,240}?IN\s*=\s*\(\s*)"S-301"', r'\1"S-201"', t)

# 4) 删 F-501 的 BLOCK 段（用下一个 "? BLOCK xxx ?" 定位）
m = re.search(r'\?\s*BLOCK\s+RSTOIC\s+"?F-501"?\s*\?', t)
if m:
    nx = re.search(r'\?\s*BLOCK\s+[A-Z0-9]+\s+"?[A-Za-z0-9-]+"?\s*\?', t[m.end():])
    end = m.end() + nx.start() if nx else m.start() + 900
    print('  删 F-501 BLOCK 段: %d 字符' % (end - m.start()), flush=True)
    t = t[:m.start()] + t[end:]

# 5) 删 S-123 的 STREAM 段
for sid in ['S-123', 'S-125', 'S-301']:
    m = re.search(r'\?\s*STREAM\s+MATERIAL\s+"?%s"?\s*\?' % re.escape(sid), t)
    if not m:
        print('  %s 无 STREAM 段' % sid, flush=True)
        continue
    nx = re.search(r'\?\s*STREAM\s+MATERIAL\s+"?S-', t[m.end():])
    if not nx:
        nx = re.search(r'\?\s*"EO-VARS"\s*\?', t[m.end():])
    end = m.end() + nx.start() if nx else min(len(t), m.start() + 700)
    print('  删 %s STREAM 段: %d 字符' % (sid, end - m.start()), flush=True)
    t = t[:m.start()] + t[end:]

# 6) 物性方法
t = re.sub(r'PARAM\s+BASE\s*=\s*"[^"]*"', 'PARAM BASE = "NRTL"', t, count=1)

open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)
print('重建 %d -> %d 字符, ? BLOCK 段数=%d' % (n0, len(t), len(re.findall(r'\?\s*BLOCK\s', t))), flush=True)

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
try:
    doc.InitFromArchive2(OUT)
except Exception as ex:
    print('打开失败:', ex, flush=True)
    sys.exit(1)
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
print('=== 3-CP 走向 ===', flush=True)
for s in ['S-106', 'S-109', 'S-110', 'S-113', 'S-114', 'S-116', 'S-118', 'S-120', 'S-121']:
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
