# -*- coding: utf-8 -*-
import re

p = r'D:\<化工工作区>\NA-Chemical-10000t_BIP.bkp'
t = open(p, encoding='utf-8', errors='ignore').read()

for bid in ['T-201', 'T-401', 'T-402', 'T-403', 'T-404']:
    m = re.search(r'\?\s*BLOCK\s+(\w+)\s+"?%s"?\s*\?' % re.escape(bid), t)
    if not m:
        print(bid, 'NOT FOUND'); continue
    m2 = re.search(r'\?\s*BLOCK\s', t[m.end():])
    end = m.end() + m2.start() if m2 else m.end() + 6000
    seg = re.sub(r'\s+', ' ', t[m.start():end])
    rrs = re.findall(r'BASIS_RR = [\d.]+', seg)
    dfs = re.findall(r'BASIS_DF = [\d.]+', seg)
    nst = re.findall(r'NSTAGE = \d+', seg)
    pres = re.findall(r'PRES1 = [\d.]+', seg)
    cond = re.findall(r'CONDENSER = \w+', seg)
    reb = re.findall(r'REBOILER = \w+', seg)
    print('%-6s %-16s %-16s %-12s %-12s %-16s %s' % (
        bid, nst[0] if nst else '-', pres[0] if pres else '-',
        rrs[0] if rrs else '-', dfs[0] if dfs else '-',
        (cond[0] if cond else '-') + '/' + (reb[0] if reb else '-'),
        m.group(1)))
