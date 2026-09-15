# -*- coding: utf-8 -*-
"""找含 WATER+TOLUENE 的 NRTL 二元参数；并 dump 一个 NRTL-1 段的真实格式"""
import sys, os, re, glob, zipfile
sys.stdout.reconfigure(encoding='utf-8')

ROOTS = [r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples',
         r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Templates',
         r'C:\Program Files\AspenTech\Aspen Plus V15.0\Engine\Examples']


def load(f):
    if f.lower().endswith('.apwz'):
        z = zipfile.ZipFile(f)
        best = ''
        for n in z.namelist():
            d = z.read(n)
            s = d.decode('latin-1', 'ignore')
            if 'NRTL' in s and len(s) > len(best):
                best = s
        return best
    return open(f, encoding='latin-1', errors='ignore').read()


files = []
for r in ROOTS:
    for ext in ('*.bkp', '*.apw', '*.apwz'):
        files += glob.glob(os.path.join(r, '**', ext), recursive=True)
files = sorted(set(files))

print('=== 含 TOLUENE+TOL 与 WATER 且带 NRTL 的示例 ===')
for f in files:
    try:
        s = load(f)
    except Exception:
        continue
    if 'NRTL' not in s:
        continue
    comp_block = s[:6000]
    has_tol = ('TOLUENE' in s[:20000]) or re.search(r'CID = "?TOL"?', s[:20000])
    has_water = ('WATER' in s[:20000]) or re.search(r'CID = "?H2O"?', s[:20000])
    if has_tol and has_water:
        print('   ', os.path.basename(f))

print()
print('=== 3phase.bkp 的 NRTL-1 段（真实格式）===')
F = None
for r in ROOTS:
    c = glob.glob(os.path.join(r, '**', '3phase.bkp'), recursive=True)
    if c:
        F = c[0]
        break
print('文件:', F)
if F:
    s = load(F)
    print('--- 组分 ---')
    for m in re.finditer(r'CID = "?([A-Z0-9-]{2,20})"?\s+ANAME = ([A-Z0-9-]+)', s[:8000]):
        print('   ', m.group(1), m.group(2))
    i = s.find('"NRTL-1"')
    print('--- NRTL-1 段 ---')
    print(repr(s[i:i+2500]))
