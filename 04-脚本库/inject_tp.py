# -*- coding: utf-8 -*-
"""把 8 条进料流股的 TEMP/PRES/BASIS/MIXED-SPEC 写进 bkp 文本段并验证"""
import time, re
import win32com.client as win32

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_模拟.bkp'
OUT = r'D:\<化工工作区>\_probe\tp_inject.bkp'

TP = {'S-101': (25, 2.0), 'S-102': (25, 2.0), 'S-103': (25, 2.0),
      'S-107': (25, 1.0), 'S-111': (40, 1.0), 'S-123': (25, 1.0),
      'S-124': (25, 1.0), 'S-125': (25, 1.0)}

t = open(SRC, encoding='utf-8', errors='ignore').read()

# 收集插入点（先定位，后统一插入，避免位移）
ins = []
for s, (T, P) in TP.items():
    m = re.search(r'\? STREAM MATERIAL "%s" \?' % re.escape(s), t)
    if not m:
        print('  未找到段 %s —— 需要新建' % s, flush=True)
        continue
    j = t.find('SUBSTREAM SSID = MIXED', m.end())
    if j < 0:
        print('  %s 段内没有 SUBSTREAM —— 跳过' % s, flush=True)
        continue
    k = j + len('SUBSTREAM SSID = MIXED')
    add = (' TEMP = %s <22> <4> PRES = %s <20> <5> BASIS = "MOLE-FLOW" MIXED-SPEC = TP'
           % (T, P))
    ins.append((k, add))

for k, add in sorted(ins, reverse=True):
    t = t[:k] + add + t[k:]
print('已插入 %d 条流股的温压' % len(ins), flush=True)
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)

# 验证
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(OUT)
time.sleep(4)
print()
print('=== 读回验证 ===', flush=True)
for s in TP:
    row = []
    for f in ['TEMP', 'PRES', 'BASIS', 'MIXED_SPEC']:
        n = doc.Tree.FindNode(r'\Data\Streams\%s\Input\%s' % (s, f))
        row.append('%s=%r' % (f, n.Value if n is not None else None))
    print('  %-8s %s' % (s, '  '.join(row)), flush=True)
try:
    doc.Close()
except Exception:
    pass
