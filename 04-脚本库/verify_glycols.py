# -*- coding: utf-8 -*-
import re, os

P = r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Datapkg\glycols.bkp'
t = open(P, encoding='utf-8', errors='ignore').read()

# 组分表
i = t.find('COMPONENTS MAIN')
seg = t[i:i+4000]
cids = re.findall(r'CID\s*=\s*("?[^"\s/]+"?)', seg)
print('glycols 组分:', sorted(set(c.strip('"') for c in cids)))

# NRTL 参数对
k = t.find('PARAMNAME = NRTL')
j = t.find('BPVAL', k)
seg2 = re.sub(r'\s+', ' ', t[k:k+60000])
pairs = set(re.findall(r'CID1\s*=\s*("?[^"\s]+"?)\s+CID2\s*=\s*("?[^"\s]+"?)', seg2))
pairs = sorted((a.strip('"'), b.strip('"')) for a, b in pairs)
print()
print('glycols NRTL 参数对总数:', len(pairs))
tol = [p for p in pairs if 'TOL' in p[0].upper() + p[1].upper()]
h2o = [p for p in pairs if 'H2O' in p[0].upper() + p[1].upper()]
print('含 TOLUENE 的对:', tol)
print('含 H2O 的对:', h2o)
print()
print('是否存在 H2O-TOLUENE 对:',
      any({a.upper(), b.upper()} == {'H2O', 'TOLUENE'} for a, b in pairs))

# BDBANK 写法
m = re.search(r'BDBANK\s*=\s*\([^)]*\)', t)
print()
print('glycols BDBANK:', re.sub(r'\s+', ' ', m.group(0)) if m else '(无 BDBANK)')
