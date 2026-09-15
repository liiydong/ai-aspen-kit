# -*- coding: utf-8 -*-
"""修正换行匹配：8 条进料流股温压全部注入"""
import time, re
import win32com.client as win32

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_模拟.bkp'
OUT = r'D:\<化工工作区>\_probe\tp_inject2.bkp'

TP = {'S-101': (25, 2.0), 'S-102': (25, 2.0), 'S-103': (25, 2.0),
      'S-107': (25, 1.0), 'S-111': (40, 1.0), 'S-123': (25, 1.0),
      'S-124': (25, 1.0), 'S-125': (25, 1.0)}

t = open(SRC, encoding='utf-8', errors='ignore').read()

# 先把已有 S-123/S-124 里插过的温压去掉，避免重复
t = re.sub(r'\s*TEMP = [\d.]+ <22> <4> PRES = [\d.]+ <20> <5> BASIS = "MOLE-FLOW" MIXED-SPEC = TP', '', t)

ins = []
for s, (T, P) in TP.items():
    m = re.search(r'\?\s*STREAM\s+MATERIAL\s+"%s"\s*\?' % re.escape(s), t)
    if not m:
        print('  ✗ 未找到段 %s' % s, flush=True)
        continue
    j = re.search(r'SUBSTREAM\s+SSID\s*=\s*MIXED', t[m.end():])
    if not j:
        print('  ✗ %s 段内无 SUBSTREAM' % s, flush=True)
        continue
    k = m.end() + j.end()
    add = (' TEMP = %s <22> <4> PRES = %s <20> <5> BASIS = "MOLE-FLOW" MIXED-SPEC = TP'
           % (T, P))
    ins.append((k, add, s))

for k, add, s in sorted(ins, reverse=True):
    t = t[:k] + add + t[k:]
print('已插入 %d 条: %s' % (len(ins), [x[2] for x in ins]), flush=True)
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(OUT)
time.sleep(4)
print()
print('=== 读回验证 ===', flush=True)
good = 0
for s in TP:
    vals = []
    for f in ['TEMP', 'PRES']:
        n = doc.Tree.FindNode(r'\Data\Streams\%s\Input\%s' % (s, f))
        vals.append(n.Value if n is not None else None)
    okk = all(v is not None for v in vals)
    good += okk
    print('  %-8s TEMP=%-8r PRES=%-6r %s' % (s, vals[0], vals[1], '✓' if okk else '✗'), flush=True)
print()
print('成功 %d / %d' % (good, len(TP)), flush=True)
try:
    doc.SaveAs(r'D:\<化工工作区>\_probe\tp_ok.bkp')
    print('已另存 tp_ok.bkp', flush=True)
except Exception as e:
    print('SaveAs 失败', str(e)[:80], flush=True)
try:
    doc.Close()
except Exception:
    pass
