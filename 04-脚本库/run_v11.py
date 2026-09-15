# -*- coding: utf-8 -*-
"""v11：按段解析安全重建 —— 去掉 NaOH/H2SO4/Na2SO4 与 F-501/S-123，物性改 NRTL"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_v7.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_v11.bkp'
REM = {'NAOH', 'H2SO4', 'NA2SO4'}

t = open(SRC, encoding='utf-8', errors='ignore').read()
print('源文件 %d 字符' % len(t), flush=True)

# ---- 按 "? XXX ?" 切段 ----
starts = [m.start() for m in re.finditer(r'(?m)^\?\s', t)]
print('段数:', len(starts), flush=True)
secs = []
for i, s in enumerate(starts):
    e = starts[i + 1] if i + 1 < len(starts) else len(t)
    secs.append(t[s:e])

def head(sec):
    m = re.match(r'\?\s*([^\n?]{0,90}?)\s*\?', sec)
    return re.sub(r'\s+', ' ', m.group(1)).strip() if m else '??'

out = []
dropped = []
for sec in secs:
    h = head(sec)
    drop = False
    if h.startswith('PROPERTIES "MOLEC-STRUCT"'):
        cid = h.split()[-1].strip('"')
        if cid in REM:
            drop = True
    if h.startswith('BLOCK RSTOIC') and '"F-501"' in h:
        drop = True
    if h.startswith('STREAM MATERIAL') and any('"%s"' % s in h for s in ['S-123', 'S-125', 'S-301']):
        drop = True
    if drop:
        dropped.append(h)
        continue
    out.append(sec)

print('删除的段:', flush=True)
for d in dropped:
    print('   -', d, flush=True)

# ---- 修改 COMPONENTS MAIN ----
for i, sec in enumerate(out):
    if head(sec).startswith('COMPONENTS MAIN'):
        m = re.match(r'(\?\s*COMPONENTS MAIN\s*\?)', sec)
        pre = m.group(1)
        body = sec[m.end():]
        ents = re.split(r'\s*/\s*', body)
        kept, gone = [], []
        for e in ents:
            mm = re.search(r'CID\s*=\s*"?([A-Za-z0-9_-]+)"?', e)
            cid = mm.group(1).strip('"') if mm else None
            (gone if cid in REM else kept).append(e)
        out[i] = pre + ' / '.join(kept)
        print('COMPONENTS MAIN 删除:', gone and [re.search(r'CID\s*=\s*"?([A-Za-z0-9_-]+)', g).group(1) for g in gone], flush=True)

# ---- 修改 FLOWSHEET ----
for i, sec in enumerate(out):
    if head(sec).startswith('FLOWSHEET'):
        s = sec
        s = re.sub(r'\\\s*BLOCK\s+BLKID\s*=\s*"F-501"[\s\S]{0,300}?OUT\s*=\s*\([^)]*\)', '', s)
        s = re.sub(r'(BLKID\s*=\s*"R-501"[\s\S]{0,240}?IN\s*=\s*\(\s*[^)]*?)\s*"S-123"\s*M0-1', r'\1', s)
        s = re.sub(r'(BLKID\s*=\s*"E-601"[\s\S]{0,240}?IN\s*=\s*\(\s*)"S-301"', r'\1"S-201"', s)
        out[i] = s
        m = re.search(r'BLKID\s*=\s*"R-501"[\s\S]{0,200}?IN\s*=\s*\([^)]*\)', s)
        print('R-501:', re.sub(r'\s+', ' ', m.group(0)) if m else '未找到', flush=True)
        m = re.search(r'BLKID\s*=\s*"E-601"[\s\S]{0,200}?IN\s*=\s*\([^)]*\)', s)
        print('E-601:', re.sub(r'\s+', ' ', m.group(0)) if m else '未找到', flush=True)
        print('F-501 是否还在 FLOWSHEET:', 'F-501' in s, flush=True)

# ---- 物性方法 ----
for i, sec in enumerate(out):
    if head(sec).startswith('PROPERTIES "OPTION-SETS"'):
        out[i] = re.sub(r'PARAM\s+BASE\s*=\s*"[^"]*"', 'PARAM BASE = "NRTL"', sec)
        print('物性方法段:', head(out[i]), '->', re.search(r'PARAM\s+BASE\s*=\s*"[^"]*"', out[i]).group(0), flush=True)

new = ''.join(out)
open(OUT, 'w', encoding='utf-8', errors='ignore').write(new)
print('重建后 %d 字符 -> %s' % (len(new), OUT), flush=True)

# ================= run =================
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
