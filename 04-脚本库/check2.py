# -*- coding: utf-8 -*-
"""扫描订正稿中残留的旧 Aspen 数值（仅论文正文 8 篇）"""
import os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
D = r'C:\Users\<用户名>\<项目原始资料>化学法版'
THESIS = ['00_封面与摘要.md', '01_第1章_绪论.md', '02_第2章_生产工艺流程设计与说明.md',
          '03_第3章_物料衡算.md', '04_第4章_热量衡算.md', '05_第5章_主要设备工艺计算与选型.md',
          '06_第6章_环境保护与安全生产.md', '07_结论参考文献与附录.md']
PATTERNS = {
 '产品流量/产量': ['1419.5', '1419.2', '10220', '10221'],
 '单耗/收率': ['910.7', '910.75', '83.73'],
 '流股流量': ['5361.4', '2666.2', '1173.7', '1492.5', '2346.2', '3492.3', '3721.5', '1375.3',
              '1280.8', '1261.5', '4221.1', '2695.2', '2173.9', '2037.3', '5805.9', '1414.1',
              '1465.0', '472.5', '1417.5', '1890.0', '1890.2', '73.0 kg/h', '93.9 kg/h'],
 '浓度': ['70.89', '35.25'],
 '塔温度/负荷': ['135.5 ℃', '54.4', '150.0', '807', '906', '818', '821', '1862 kW', '1755 kW', '3205 kW'],
 '蒸汽/热': ['6.63 MW', '11.8 t/h', '23880 MJ/h', '933.8', '6.72×10⁶', '42200'],
 '其他': ['247.5 t/a', '88.4 t', '5189', 'T-203', 'DN1100'],
}
for fn in THESIS:
    p = os.path.join(D, fn)
    if not os.path.exists(p):
        print('=== %s MISSING' % fn); continue
    hits = []
    for i, ln in enumerate(open(p, encoding='utf-8').read().split('\n'), 1):
        for grp, pats in PATTERNS.items():
            for o in pats:
                if o in ln:
                    hits.append('  L%-5d [%s|%s] %s' % (i, grp, o, ln.strip()[:130]))
                    break
            else:
                continue
            break
    print('=== %s : %d hit(s)' % (fn, len(hits)))
    for h in hits:
        print(h)
