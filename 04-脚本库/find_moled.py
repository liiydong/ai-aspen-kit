# -*- coding: utf-8 -*-
"""找 COL-SPECS 中固定采出量的写法（MOLE-D / MASS-D）"""
import sys, os, re, glob, zipfile
sys.stdout.reconfigure(encoding='utf-8')

ROOTS = [r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples',
         r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Templates']


def load(f):
    if f.lower().endswith('.apwz'):
        z = zipfile.ZipFile(f)
        best = ''
        for n in z.namelist():
            s = z.read(n).decode('latin-1', 'ignore')
            if 'MOLE-D' in s and len(s) > len(best):
                best = s
        return best
    return open(f, encoding='latin-1', errors='ignore').read()


files = []
for r in ROOTS:
    for ext in ('*.bkp', '*.apwz'):
        files += glob.glob(os.path.join(r, '**', ext), recursive=True)
n = 0
for f in sorted(set(files)):
    try:
        s = load(f)
    except Exception:
        continue
    for m in re.finditer(r'MOLE-D = [^\n\\]{0,40}', s):
        print('%-46s %s' % (os.path.basename(f)[:45], m.group()))
        n += 1
        break
    if n >= 12:
        break
print('共', n)
