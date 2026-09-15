# -*- coding: utf-8 -*-
"""对照实验：FREE_WATER / NPHASE / 亨利组分 三种处理对警告数的影响"""
import sys, re, time, os, json
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_BIP2.bkp'
OUT = r'D:\<化工工作区>\_probe'
t0 = open(SRC, encoding='utf-8', errors='ignore').read()

# ---------- 变体构造 ----------
VAR = {}
VAR['A_base'] = t0
VAR['B_freewater'] = t0.replace('FREE_WATER  (n=0) val=NO', 'FREE_WATER  (n=0) val=YES')  # 占位，下面用真替换
# 真替换（文本形式）
m = re.search(r'FREE_WATER\s*=\s*NO', t0)
print('FREE_WATER 文本定位:', bool(m), m.group(0) if m else '')
VAR['B_freewater'] = t0.replace(m.group(0), 'FREE_WATER = YES', 1) if m else t0

m2 = re.search(r'NPHASE\s*=\s*2', t0)
print('NPHASE 文本定位:', bool(m2), m2.group(0) if m2 else '')
VAR['C_nphase3'] = t0.replace(m2.group(0), 'NPHASE = 3', 1) if m2 else t0

# 亨利组分段
m3 = re.search(r'\?\s*COMPONENTS\s+"COMP-LIST"', t0)
if not m3:
    m3 = re.search(r'\?\s*COMPONENTS\s+MAIN\s*\?', t0)
    # 找 MAIN 段结尾
    end = t0.find('\n? ', m3.end())
print('insert 锚点:', bool(m3))
HENRY = ('? COMPONENTS "HENRY-COMPS" MAIN ?\n'
         '\\ "HENRY-COMPS" CID = ( N2 O2 CO CO2 HCN ) \\\n')
VAR['D_henry'] = t0[:m3.start()] + HENRY + t0[m3.start():]
# E: NPHASE3 + 亨利
if m2:
    tmp = t0.replace(m2.group(0), 'NPHASE = 3', 1)
    VAR['E_np3_henry'] = tmp[:m3.start()] + HENRY + tmp[m3.start():]


def run(src, tag):
    MSGS = []

    class Sink:
        def OnControlPanelMessage(self, *a):
            s = ' '.join(str(x) for x in a).strip()
            if s and s != 'False':
                MSGS.append(s)

    doc = win32.DispatchEx('Apwn.Document')
    doc.SuppressDialogs = True
    win32.WithEvents(doc, Sink)
    try:
        doc.InitFromArchive2(src)
    except Exception as e:
        print(tag, 'LOAD ERR', str(e)[:70]); return
    time.sleep(2)
    doc.Engine.Run2(False)
    t1 = time.time()
    while time.time() - t1 < 600:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.25)
        if time.time() - t1 > 10:
            try:
                if not bool(doc.Engine.IsRunning):
                    break
            except Exception:
                break
    for _ in range(70):
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)

    def g(p):
        n = doc.Tree.FindNode(p)
        try:
            return n.Value if n is not None else None
        except Exception:
            return None

    warn = [x for x in MSGS if re.search(r'\*\s*WARNING', x)]
    terr = [x for x in MSGS if 'Terminal Errors' in x]
    IN = ['S-101', 'S-102', 'S-103', 'S-107', 'S-111', 'S-124', 'S-132', 'S-210']
    OUTL = ['S-130', 'S-131', 'S-211', 'S-212', 'S-213', 'S-115', 'S-119',
            'S-122', 'S-216', 'S-217', 'S-218', 'S-405C', 'S-406', 'S-209']
    vi = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in IN)
    vo = sum(float(g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s) or 0) for s in OUTL)
    prod = g(r'\Data\Streams\S-209\Output\MASSFLMX\MIXED')
    print('%-14s 警告=%-3d %s  进=%.1f 出=%.1f 偏差=%.4f%%  产品=%.2f' % (
        tag, len(warn), (terr[0].replace('False', '').strip()[:40] if terr else ''),
        vi, vo, abs(vi - vo) / vi * 100 if vi else 0, float(prod or 0)))
    kinds = set()
    for w in warn:
        kinds.add(re.sub(r'\s+', ' ', w.replace('False', '').strip())[:26])
    for k in sorted(kinds):
        print('       ·', k)
    try:
        doc.Close()
    except Exception:
        pass


print()
for tag in ['A_base', 'B_freewater', 'C_nphase3', 'D_henry', 'E_np3_henry']:
    src = os.path.join(OUT, 'wv_%s.bkp' % tag)
    open(src, 'w', encoding='utf-8', errors='ignore').write(VAR[tag])
    run(src, tag)
print('DONE')
