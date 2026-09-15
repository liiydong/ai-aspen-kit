# -*- coding: utf-8 -*-
"""跨章一致性复查：找出各章之间互相矛盾的数字"""
import os, re, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
D = r'C:\Users\<用户名>\<项目原始资料>化学法版'
THESIS = ['00_封面与摘要.md', '01_第1章_绪论.md', '02_第2章_生产工艺流程设计与说明.md',
          '03_第3章_物料衡算.md', '04_第4章_热量衡算.md', '05_第5章_主要设备工艺计算与选型.md',
          '06_第6章_环境保护与安全生产.md', '07_结论参考文献与附录.md']
# 关键词：同一物理量的不同写法，应当一致
GROUPS = {
 '公用工程总费用(万元)': [r'1[45]\d\d 万元', r'公用工程年费用?约?\s*\*{0,2}(1[45]\d\d)'],
 '蒸汽年消耗': [r'4[0-9]{4} t/a'],
 '循环水年消耗': [r'[0-9.]+×10⁶ t/a', r'[0-9.]+×10⁶ t'],
 '蒸汽小时需求': [r'[0-9.]+ t/h'],
 '熔盐产汽': [r'3\.9 t/h', r'11\.8 t/h'],
 '蒸汽缺口/富余': [r'缺口[^，。]*', r'富余[^，。]*', r'自给'],
 '反应器热': [r'2\.19 MW', r'6\.63 MW', r'2193\.6 kW'],
 '产品产量': [r'1016[34](\.\d)?', r'1022[01]'],
 '循环水总量': [r'3\.93×10⁶', r'4\.81×10⁶', r'6\.72×10⁶'],
 '公用工程占比': [r'占生产成本的\s*[0-9.]+%'],
}
for g, pats in GROUPS.items():
    print('=== %s ===' % g)
    seen = {}
    for fn in THESIS:
        t = open(os.path.join(D, fn), encoding='utf-8').read()
        for i, ln in enumerate(t.split('\n'), 1):
            for p in pats:
                for m in re.finditer(p, ln):
                    key = m.group(0)
                    seen.setdefault(key, []).append('%s:L%d' % (fn[:6], i))
    for k in sorted(seen):
        locs = seen[k]
        print('   %-22s x%-3d  %s' % (k, len(locs), ', '.join(locs[:5])))
    print()
