# -*- coding: utf-8 -*-
import sys, re, os
sys.stdout.reconfigure(encoding='utf-8')
p = r'D:\<化工工作区>\_probe\t403_radfrac.bkp'
L = []
L.append('exists=%s size=%s' % (os.path.exists(p), os.path.getsize(p) if os.path.exists(p) else 0))
t = open(p, encoding='utf-8', errors='ignore').read()
L.append('size read = %d' % len(t))
i = t.find('T-403')
L.append('--- 注册表附近 ---')
L.append(repr(t[max(0, i - 200):i + 80]))
L.append('')
L.append('--- FLOWSHEET T-403 ---')
m = re.search(r'.{0,80}T-403.{0,220}', t[t.find('FLOWSHEET GLOBAL'):])
L.append(m.group() if m else 'not found')
L.append('')
L.append('--- 段头 ---')
for m in re.finditer(r'\?\s*BLOCK\s*\n?\s*([A-Z]+)\s*\n?\s*"T-403"', t):
    L.append(repr(t[m.start():m.start() + 120]))
L.append('')
L.append('--- DSET BLOCK ... T-403 计数 ---')
L.append('SEP=%d RADFRAC=%d' % (len(re.findall(r'BLOCK SEP T-403', t)), len(re.findall(r'BLOCK RADFRAC T-403', t))))
open(r'D:\<化工工作区>\_probe\t403file.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('DONE')
