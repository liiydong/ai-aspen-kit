# -*- coding: utf-8 -*-
"""修正版：按记录边界切分，收集本体系所有真实 NRTL 二元参数"""
import glob, os, re, json

FILES = glob.glob(r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\**\*.bkp', recursive=True)
FILES += glob.glob(r'C:\ProgramData\AspenTech\APED V15.0\**\*.bkp', recursive=True)

ALIAS = {
    '3-MP': ['3-METHYLPYRIDINE', 'C6H7N-D2'],
    '4-MP': ['4-METHYLPYRIDINE', 'C6H7N-2'],
    '3-CP': ['NICOTINONITRILE', 'C6H4N2', '3-CYANOPYRIDINE'],
    '4-CP': ['4-PYRIDINENITRILE', 'C6H4N2-N1', '4-CYANOPYRIDINE'],
    'NAM': ['NICOTINIC-ACID-AMIDE', 'C6H6N2O'],
    'NAC': ['NIACIN', 'C6H5NO2-D1'],
    'H2O': ['H2O', 'WATER'],
    'TOL': ['TOLUENE', 'C7H8'],
    'NH3': ['NH3', 'AMMONIA', 'H3N'],
    'CO2': ['CO2', 'CARBON-DIOXIDE'],
    'HCN': ['HCN', 'HYDROGEN-CYANIDE', 'CHN'],
    'CO': ['CO', 'CARBON-MONOXIDE'],
    'O2': ['O2', 'OXYGEN'],
    'N2': ['N2', 'NITROGEN'],
}
REV = {v.upper(): k for k, vs in ALIAS.items() for v in vs}

found = {}
for p in FILES:
    try:
        t = open(p, encoding='utf-8', errors='ignore').read()
    except Exception:
        continue
    if 'PARAMNAME2 = NRTL' not in t:
        continue
    flat = re.sub(r'\s+', ' ', t)
    ms = list(re.finditer(r'PARAMNAME2 = NRTL\s+CID1 = ("?[^"\s]+"?)\s+CID2 = ("?[^"\s]+"?)', flat))
    for k, m in enumerate(ms):
        end = ms[k + 1].start() if k + 1 < len(ms) else min(len(flat), m.start() + 1200)
        rec = flat[m.start():end]
        if 'UVAL' not in rec:
            continue
        a = REV.get(m.group(1).strip('"').upper())
        b = REV.get(m.group(2).strip('"').upper())
        if not a or not b or a == b:
            continue
        key = tuple(sorted([a, b]))
        if key not in found:
            found[key] = (rec[:900], os.path.basename(p))

print('本体系可用的真实参数对:', len(found))
out = {}
for k, (rec, f) in sorted(found.items()):
    print()
    print('=== %s / %s  [%s] ===' % (k[0], k[1], f))
    print(rec)
    out['%s-%s' % k] = {'src': f, 'rec': rec}

open(r'D:\<化工工作区>\_probe\pairs_real.json', 'w', encoding='utf-8').write(json.dumps(out, ensure_ascii=False, indent=1))
print()
print('已存 pairs_real.json')
