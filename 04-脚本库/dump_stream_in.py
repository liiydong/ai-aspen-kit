# -*- coding: utf-8 -*-
"""找 S-111 等进料流股的流量定义段"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')

t = open(r'D:\<化工工作区>\_probe\bipA2.bkp', encoding='utf-8', errors='ignore').read()
# 找所有 '? STREAM' 段
ms = list(re.finditer(r'\?\s*STREAM\s+([A-Z ]*?)\s*"([A-Za-z0-9-]+)"\s*\?', t))
print('STREAM 段 %d 个:' % len(ms))
for m in ms:
    print('  @%-8d %s' % (m.start(), m.group(0).replace('\n', ' ')))
print()
i = t.find('? STREAM')
print('=== 第一个 STREAM 段全文（900 字符）===')
print(t[i:i + 900])
print()
# 看 S-111 的输入在哪
i2 = t.find('S-111')
while i2 > 0 and i2 < 200000:
    seg = t[max(0, i2 - 120):i2 + 160]
    if 'DSET' not in seg:
        print('@%-8d %s' % (i2, repr(seg)))
    i2 = t.find('S-111', i2 + 1)
