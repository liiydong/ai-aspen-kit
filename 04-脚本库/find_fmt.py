# -*- coding: utf-8 -*-
"""取齐各类模块的 bkp 段落格式（COMPR / PUMP / MIXER / FSPLIT / HEATER / FLASH2 / VALVE / CFUGE）"""
import sys, os, re, glob, zipfile
sys.stdout.reconfigure(encoding='utf-8')

ROOTS = [r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples',
         r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Templates']
WANT = ['COMPR', 'PUMP', 'MIXER', 'FSPLIT', 'HEATER', 'FLASH2', 'VALVE', 'CFUGE', 'CRYSTALLIZER', 'DRYER']


def load(f):
    if f.lower().endswith('.apwz'):
        z = zipfile.ZipFile(f)
        best = ''
        for n in z.namelist():
            s = z.read(n).decode('latin-1', 'ignore')
            if '? BLOCK' in s and len(s) > len(best):
                best = s
        return best
    return open(f, encoding='latin-1', errors='ignore').read()


files = []
for r in ROOTS:
    for ext in ('*.bkp', '*.apwz'):
        files += glob.glob(os.path.join(r, '**', ext), recursive=True)
files = sorted(set(files))
print('示例文件数:', len(files))

found = {}
for f in files:
    if len(found) >= len(WANT):
        break
    try:
        s = load(f)
    except Exception:
        continue
    for w in WANT:
        if w in found:
            continue
        m = re.search(r'\?\s*BLOCK\s+' + w + r'\s+"?[A-Za-z0-9_.-]+"?\s*\?', s)
        if not m:
            continue
        j = s.find('? BLOCK', m.end())
        blk = s[m.start():j if j > 0 else m.start() + 900]
        found[w] = (os.path.basename(f), re.sub(r'\s+', ' ', blk)[:600])

for w in WANT:
    print('=' * 78)
    if w in found:
        f, b = found[w]
        print('%s  <- %s' % (w, f))
        print('   ', b)
    else:
        print('%s  未找到' % w)
