# -*- coding: utf-8 -*-
"""提取 Aspen 报告 BLOCK SECTION 的关键结果。"""
import re, sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

SRC = r'D:\<化工工作区>\_probe\report_cols.tsv'
OUT = r'D:\<化工工作区>\_probe\report_blocks.txt'

raw = []
for ln in open(SRC, encoding='utf-8').read().split('\n'):
    if not ln.strip():
        continue
    p = ln.split('\t')
    r = int(p[0]); cols = (p[1:] + ['', '', ''])[:3]
    a, b, c = cols[0], cols[1], cols[2]
    if b and c and not b.endswith(('-', '+')):
        b = b + ' '
    raw.append((r, (a + b + c).rstrip()))

BLK = re.compile(r'^BLOCK:\s+(\S+)\s+MODEL:\s+(\S+)')
KEYS = [
    'NUMBER OF STAGES', 'MOLAR REFLUX RATIO', 'DISTILLATE TO FEED RATIO',
    'MOLAR VAPOR DIST / TOTAL DIST', 'TOP STAGE TEMPERATURE', 'BOTTOM STAGE TEMPERATURE',
    'TOP STAGE PRESSURE', 'TOP STAGE LIQUID FLOW', 'BOTTOM STAGE LIQUID FLOW',
    'BOILUP VAPOR FLOW', 'CONDENSER DUTY', 'REBOILER DUTY', 'MOLAR BOILUP RATIO',
    'CONDENSER PRESSURE', 'REBOILER PRESSURE', 'FEED STAGE', 'STAGE 1 PRES',
    'TEMP, C', 'PRES, BAR', 'CALCULATED MINIMUM REFLUX RATIO',
]
blk = None
blocks = {}
order = []
cur = None
for r, s in raw:
    st = s.strip()
    m = BLK.match(st.replace('  ', ' '))
    if m and '.' not in m.group(2):
        name = m.group(1)
        if name != cur:
            cur = name
            blocks[name] = {'model': m.group(2), 'lines': []}
            order.append(name)
        continue
    if cur:
        blocks[cur]['lines'].append(st)

out = []
for name in order:
    b = blocks[name]
    out.append('=' * 90)
    out.append('BLOCK %s   MODEL %s' % (name, b['model']))
    for ln in b['lines']:
        u = ln.strip()
        if not u:
            continue
        if any(u.startswith(k) for k in KEYS) or u.startswith(('INLETS', 'OUTLETS', 'S-', 'MOLE(', 'MASS(', 'ENTHALPY(')) \
           or re.match(r'^\d+\s', u) and ('STAGE' in u or 'TEMP' in u) \
           or 'TOP STAGE' in u or 'BOTTOM STAGE' in u or 'DUTY' in u or 'REFLUX' in u or 'STAGES' in u \
           or 'BLOCK TYPE' in u or re.match(r'^[A-Z0-9\-]{3,12}\s{2,}[-+]?\d', u) and len(u) < 80:
            if 'COMPONENT' in u and 'SPLIT' in u:
                continue
            out.append('   ' + u[:150])

open(OUT, 'w', encoding='utf-8').write('\n'.join(out))
print('blocks:', len(order))
print(', '.join(order))
