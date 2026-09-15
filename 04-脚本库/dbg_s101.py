# -*- coding: utf-8 -*-
"""debug：跟踪 S-101 页每行的解析"""
import re, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

COMPONENTS = ['3-MP','NH3','O2','N2','3-CP','4-CP','H2O','TOL','CO2','HCN','CO','AIR','NAM','NAC','4-MP']
NUM = re.compile(r'[-+]?(?:\d+\.\d+|\d+)(?:[-+]\d+)?')
STREAMID = re.compile(r'S-\d+[A-Z]?')
PAGEHDR = re.compile(r'^(?:S-\d+[A-Z]?\s+)+S-\d+[A-Z]?$')

raw = []
for ln in open(r'D:\<化工工作区>\_probe\report_cols.tsv', encoding='utf-8').read().split('\n'):
    if not ln.strip():
        continue
    p = ln.split('\t')
    raw.append((int(p[0]), ''.join(p[1:]).rstrip()))

active = False
n = 5
out = []
for r, s in raw:
    st = s.strip()
    if PAGEHDR.match(st):
        ids = STREAMID.findall(st)
        active = ('S-101' in ids)
        n = len(ids)
        out.append('### R%d PAGE %s n=%d' % (r, ids, n))
        continue
    if not active:
        continue
    if st.startswith('STREAM SECTION') or st.startswith('U-O-S'):
        active = False
        out.append('### END at R%d' % r)
        break
    hit = None
    for comp in COMPONENTS:
        if st.startswith(comp) and (len(st) == len(comp) or not st[len(comp)].isalnum()):
            toks = NUM.findall(s)
            out.append('R%-5d comp=%-5s ntok=%d toks=%s || RAW=%r' % (r, comp, len(toks), toks[-n:] if len(toks) >= n else toks, s))
            hit = comp
            break
    if hit is None and st.startswith(('TEMP','PRES','VFRAC','KCAL','CAL/','KMOL','KG/CUM','CUM/HR','AVG','LFRAC','GCAL')):
        toks = NUM.findall(s)
        out.append('R%-5d tag =%-5s ntok=%d toks=%s' % (r, st.split()[0], len(toks), toks[-n:] if len(toks) >= n else toks))
print('\n'.join(out))
