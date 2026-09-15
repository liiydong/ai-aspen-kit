# -*- coding: utf-8 -*-
"""诊断 v13 文件里的 DATABANKS / 二元参数 / 物性方法状态"""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')

P = r'D:\<化工工作区>\NA-Chemical-10000t_v13_s.bkp'
t = open(P, encoding='utf-8', errors='ignore').read()
print('文件大小:', len(t))

# 1. 段名清单
print('\n=== 段名清单 ===')
for m in re.finditer(r'\?\s*([^?\\]{1,45}?)\s*\?', t):
    s = m.group(1).strip()
    if s and len(s) < 46:
        pass
secs = [m.group(1).strip() for m in re.finditer(r'\?\s*([^?\\]{1,45}?)\s*\?', t)]
from collections import Counter
c = Counter(secs)
for k, v in c.most_common(50):
    print('  %3d  %s' % (v, k))

# 2. DATABANKS 段
print('\n=== DATABANKS 段 ===')
i = t.find('? DATABANKS ?')
if i >= 0:
    print(repr(t[i:i+400]))
else:
    print('  未找到 DATABANKS 段')

# 3. NRTL 相关
print('\n=== NRTL 出现位置 ===')
for m in list(re.finditer(r'.{0,60}NRTL.{0,80}', t))[:8]:
    print('  ', re.sub(r'\s+', ' ', m.group()))
print('  NRTL 总出现次数:', t.count('NRTL'))

# 4. BINARY / BIP
print('\n=== BINARY 相关 ===')
for kw in ['BINARY', 'BIP', 'LLE-ASPEN', 'VLE-IG', 'VLE-RK', 'NRTL-2', 'ESTIMATE']:
    print('  %-12s %d' % (kw, t.count(kw)))

# 5. PROPERTIES MAIN
print('\n=== PROPERTIES MAIN 段 ===')
i = t.find('? PROPERTIES MAIN ?')
print(repr(re.sub(r'\s+', ' ', t[i:i+500])) if i >= 0 else '未找到')

# 6. 是否有 OPTION-SETS
print('\n=== OPTION-SETS ===')
i = t.find('OPTION-SETS')
print(repr(re.sub(r'\s+', ' ', t[i-100:i+500])) if i >= 0 else '未找到')
