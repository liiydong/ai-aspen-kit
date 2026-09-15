# -*- coding: utf-8 -*-
import json, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
d = json.load(open(r'D:\<化工工作区>\_probe\report_parsed.json', encoding='utf-8'))
S = d['streams']
COMP = ['3-MP','NH3','O2','N2','3-CP','4-CP','H2O','TOL','CO2','HCN','CO','AIR','NAM','NAC','4-MP']
out = []
out.append('total streams: %d' % len(S))
for sid, v in S.items():
    c = v.get('comp_kmol', {})
    miss = [x for x in COMP if x not in c]
    out.append('%-8s n=%2d miss=%s' % (sid, len(c), ','.join(miss)))
out.append('')
out.append('S-101: ' + json.dumps(S.get('S-101', {}).get('comp_kmol', {}), ensure_ascii=False))
out.append('S-109: ' + json.dumps(S.get('S-109', {}).get('comp_kmol', {}), ensure_ascii=False))
print('\n'.join(out))
