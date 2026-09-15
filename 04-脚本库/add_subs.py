# -*- coding: utf-8 -*-
"""补 ? "STREAM-CLASS" SUBSTREAMS ? 段，再运行"""
import time, re, os
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\noair2.bkp'
OUT = r'D:\<化工工作区>\_probe\subs1.bkp'

SEC = ('? "STREAM-CLASS" SUBSTREAMS ? \n'
       '\\ DEF-SUBS-CLA SUBSTREAMS = MIXED \n'
       '\\ \\ DEF-SUBS-CLA SUBSTREAMS = CISOLID \n'
       '\\ \\ DEF-SUBS-CLA SUBSTREAMS = NC \n'
       '\\ \\ DEF-SUBS-CLA SUBSTREAMS = NCPSD \n'
       '\\ \\ DEF-SUBS-CLA SUBSTREAMS = CIPSD \\ \n')

t = open(SRC, encoding='utf-8', errors='ignore').read()
if 'STREAM-CLASS" SUBSTREAMS' in t:
    print('已有该段，跳过注入', flush=True)
    t2 = t
else:
    m = re.search(r'\?\s*"PROP-SET"\s+MAIN\s+HXDESIGN\s*\?', t)
    if not m:
        print('未找到插入锚点!', flush=True)
        raise SystemExit(1)
    t2 = t[:m.start()] + SEC + t[m.start():]
    print('已插入 SUBSTREAMS 段 (位置 %d)' % m.start(), flush=True)
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t2)

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(OUT)
time.sleep(4)
print('打开后 \\Data 子节点数:', doc.Tree.FindNode(r'\Data').Elements.Count, flush=True)
print('打开 Ready =', doc.Engine.Ready, flush=True)

t0 = time.time()
try:
    doc.Engine.Run2()
except Exception as e:
    print('Run2 异常', str(e)[:120], flush=True)
while time.time() - t0 < 150:
    time.sleep(3)
    try:
        if doc.Engine.IsRunning is False:
            break
    except Exception:
        break
print('运行用时 %.0fs' % (time.time() - t0), flush=True)
print()
got = []
for b in ['M-101', 'E-101', 'R-101', 'E-104', 'T-201', 'T-301', 'T-401', 'T-404',
          'R-501', 'F-501', 'E-601']:
    n = doc.Tree.FindNode(r'\Data\Blocks\%s\Output' % b)
    if n is None:
        got.append('%s:无' % b)
        continue
    vals = []
    for f in ['TEMP', 'PRES', 'DUTY', 'MASS_RR', 'MOLE_RR']:
        try:
            x = n.Elements.Item(f)
            if x is not None and x.Value not in (None, ''):
                vals.append('%s=%s' % (f, x.Value))
        except Exception:
            pass
    got.append('%s[%s]' % (b, ','.join(vals) if vals else '空'))
print('  ', ' '.join(got), flush=True)
print()
for s in ['S-104', 'S-105', 'S-106', 'S-110', 'S-201', 'S-401']:
    n = doc.Tree.FindNode(r'\Data\Streams\%s\Output\TEMP' % s)
    print('   %s.Out.TEMP = %r' % (s, n.Value if n is not None else None), flush=True)
try:
    doc.SaveAs(r'D:\<化工工作区>\_probe\subs1_saved.bkp')
    print('  已保存', flush=True)
except Exception as e:
    print('  SaveAs 失败', str(e)[:80], flush=True)
try:
    doc.Close()
except Exception:
    pass
