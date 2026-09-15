# -*- coding: utf-8 -*-
"""从 Aspen 出厂示例中提取 H2O-TOLUENE 的真实 NRTL 参数记录"""
import re

P = r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Datapkg\glycols.bkp'
t = open(P, encoding='utf-8', errors='ignore').read()

i = t.find('PARAMNAME = NRTL')
j = t.find('PARAMNAME =', i + 20)
seg = t[i:j] if j > i else t[i:i+60000]
flat = re.sub(r'\s+', ' ', seg)

# 切分出每条 BPVAL 记录
recs = re.split(r'BPVAL', flat)
for r in recs:
    if 'CID1 = H2O' in r and 'CID2 = TOLUENE' in r:
        print('=== H2O-TOLUENE 记录 ===')
        print(r[:900])
        print()
    if 'CID1 = TOLUENE' in r and 'CID2 = H2O' in r:
        print('=== TOLUENE-H2O 记录 ===')
        print(r[:900])
        print()

# 段头
print('=== NRTL 段头 ===')
print(flat[:230])
