# -*- coding: utf-8 -*-
"""扫描订正稿中残留的旧 Aspen 数值"""
import os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
D = r'C:\Users\<用户名>\<项目原始资料>化学法版'
OLDS = ['1419.5', '1419.2', '1419.0', '10220', '10221', '910.7', '910.75', '83.73',
        '5361.4', '2666.2', '1173.7', '1492.5', '2346.2', '3492.3', '3721.5', '1375.3',
        '1280.8', '1261.5', '4221.1', '2695.2', '2173.9', '2037.3', '5805.9', '1414.1',
        '1465.0', '472.5', '1417.5', '70.89', '35.25', '135.5 ℃', '34.3 ', '54.4',
        '6.63 MW', '11.8 t/h', '1862 kW', '1755 kW', '3205 kW', '1419.256', '1419.529',
        '2321.4', '1352.8', '1079.8', '93.9 kg/h', '1890.0', '1890.2', '73.0 kg/h']
for fn in sorted(os.listdir(D)):
    if not fn.endswith('.md'):
        continue
    p = os.path.join(D, fn)
    hits = []
    for i, ln in enumerate(open(p, encoding='utf-8').read().split('\n'), 1):
        for o in OLDS:
            if o in ln:
                hits.append('  L%-5d [%s] %s' % (i, o, ln.strip()[:150]))
                break
    print('=== %s : %d hit(s)' % (fn, len(hits)))
    for h in hits:
        print(h)
