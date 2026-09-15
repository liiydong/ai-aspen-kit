# -*- coding: utf-8 -*-
"""输出流股汇总表（紧凑），供与论文逐项核对。"""
import json, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
d = json.load(open(r'D:\<化工工作区>\_probe\report_parsed.json', encoding='utf-8'))
S = d['streams']
COMP = ['3-MP','NH3','O2','N2','3-CP','4-CP','H2O','TOL','CO2','HCN','CO','AIR','NAM','NAC','4-MP']
MW = {'3-MP':93.13,'NH3':17.03,'O2':32.0,'N2':28.01,'3-CP':104.11,'4-CP':104.11,'H2O':18.02,
      'TOL':92.14,'CO2':44.01,'HCN':27.03,'CO':28.01,'AIR':28.85,'NAM':122.13,'NAC':123.11,'4-MP':93.13}

order = sorted(S.keys(), key=lambda x: (x[:2] != 'S-', int(''.join(ch for ch in x.split('-')[1] if ch.isdigit())), x))
out = []
out.append('%-9s %11s %11s %8s %7s %7s  %s' % ('STREAM','KMOL/HR','KG/HR','T(C)','P(bar)','VF','主要组成(mol%)'))
for sid in order:
    v = S[sid]
    tot = v.get('F_KMOL') or 0
    comp = v.get('comp_kmol', {})
    tops = sorted(((comp.get(c,0), c) for c in COMP), reverse=True)[:4]
    s = ', '.join('%s %.2f' % (c, (x/tot*100 if tot else 0)) for x, c in tops if x)
    out.append('%-9s %11.4f %11.2f %8.3f %7.4f %7.3f  %s' % (
        sid, tot, v.get('F_KG', 0), v.get('TEMP', float('nan')), v.get('PRES', float('nan')),
        v.get('VFRAC', float('nan')), s))
print('\n'.join(out))
