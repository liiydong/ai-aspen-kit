# -*- coding: utf-8 -*-
"""在 Aspen 自带示例里找 BLOCK SEP 的段落写法"""
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
            s = z.read(n).decode('latin-1', 'ignore')
            if 'BLOCK SEP' in s and len(s) > len(best):
                best = s
        return best
    return open(f, encoding='latin-1', errors='ignore').read()


files = []
for r in ROOTS:
    for ext in ('*.bkp', '*.apw', '*.apwz'):
        files += glob.glob(os.path.join(r, '**', ext), recursive=True)
files = sorted(set(files))
found = 0
for f in files:
    try:
        s = load(f)
    except Exception:
        continue
    i = s.find('BLOCK SEP')
    if i < 0:
        continue
    print('=' * 74)
    print(os.path.basename(f))
    # 找段头，往前退到 '? BLOCK'
    j = s.rfind('? BLOCK', 0, i)
    k = s.find('? BLOCK', i + 10)
    print(s[j:k if k > 0 else i + 1500][:1500])
    print()
    found += 1
    if found >= 3:
        break
print('找到', found, '个')
