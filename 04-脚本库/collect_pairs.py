# -*- coding: utf-8 -*-
"""扫描全部 Aspen 示例文件，收集属于本模型组分体系的真实 NRTL 二元参数记录"""
import glob, os, re

FILES = glob.glob(r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\**\*.bkp', recursive=True)
FILES += glob.glob(r'C:\ProgramData\AspenTech\APED V15.0\**\*.bkp', recursive=True)
print('文件数', len(FILES))

ALIAS = {
    '3-MP': ['3-METHYLPYRIDINE', 'C6H7N-D2', 'METHYLPYRIDINE'],
    '4-MP': ['4-METHYLPYRIDINE', 'C6H7N-2', '4-METHYLPYRIDINE-2'],
    '3-CP': ['NICOTINONITRILE', 'C6H4N2', '3-CYANOPYRIDINE'],
    '4-CP': ['4-PYRIDINENITRILE', 'C6H4N2-N1', '4-CYANOPYRIDINE'],
    'NAM': ['NICOTINIC-ACID-AMIDE', 'C6H6N2O', 'NICOTINAMIDE'],
    'NAC': ['NIACIN', 'C6H5NO2-D1', 'NICOTINIC-ACID'],
    'H2O': ['H2O', 'WATER'],
    'TOL': ['TOLUENE', 'C7H8'],
    'NH3': ['NH3', 'AMMONIA', 'H3N'],
    'CO2': ['CO2', 'CARBON-DIOXIDE'],
    'HCN': ['HCN', 'HYDROGEN-CYANIDE', 'CHN'],
    'CO': ['CO', 'CARBON-MONOXIDE'],
    'O2': ['O2', 'OXYGEN'],
    'N2': ['N2', 'NITROGEN'],
}
REV = {}
for k, vs in ALIAS.items():
    for v in vs:
        REV[v.upper()] = k


def norm(n):
    n = n.strip('"').upper()
    return REV.get(n)


found = {}   # (a,b) -> (record_text, source_file)
for p in FILES:
    try:
        t = open(p, encoding='utf-8', errors='ignore').read()
    except Exception:
        continue
    if 'PARAMNAME2 = NRTL' not in t:
        continue
    flat = re.sub(r'\s+', ' ', t)
    for m in re.finditer(r'BPVAL PARAMNAME2 = NRTL (.*?)(?=BPVAL |/ PARAMNAME2|\\ \? PROPERTIES|$)', flat):
        rec = m.group(1)
        c1 = re.search(r'CID1 = ("?[^"\s]+"?)', rec)
        c2 = re.search(r'CID2 = ("?[^"\s]+"?)', rec)
        if not (c1 and c2):
            continue
        a, b = norm(c1.group(1)), norm(c2.group(1))
        if not a or not b:
            continue
        if 'UVAL' not in rec:
            continue
        key = tuple(sorted([a, b]))
        if key not in found:
            found[key] = (rec, os.path.basename(p))

print()
print('匹配到本体系组分的真实参数对:', len(found))
for k, (rec, f) in sorted(found.items()):
    print()
    print('=== %s / %s   [来自 %s] ===' % (k[0], k[1], f))
    print(rec[:700])
