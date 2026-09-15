import glob, os, re, sys
sys.stdout.reconfigure(encoding='utf-8')

roots = [r'C:/Program Files/AspenTech/Aspen Plus V15.0',
         r'C:/Program Files/Common Files/AspenTech Shared',
         r'C:/ProgramData/AspenTech']
hits = []
for r in roots:
    for p in glob.glob(os.path.join(r, '**', '*.bkp'), recursive=True):
        try:
            t = open(p, encoding='utf-8', errors='ignore').read(2000000)
        except Exception:
            continue
        for m in re.finditer(r'BLOCK\s+RADFRAC\s+\S+', t):
            nxt = re.search(r'\n\?\s*BLOCK\s', t[m.start() + 10:])
            end = m.start() + 10 + nxt.start() if nxt else min(len(t), m.start() + 3000)
            seg = t[m.start():end]
            flat = re.sub(r'\s+', ' ', seg)
            if 'CONDENSER = NONE' in flat and 'REBOILER = NONE' in flat:
                hits.append((os.path.basename(p), flat[:1200]))

print('无冷凝器+无再沸器 的 RadFrac 数:', len(hits))
seen = set()
for name, flat in hits[:40]:
    key = flat[:400]
    if key in seen:
        continue
    seen.add(key)
    print('=' * 60)
    print('#', name)
    print(flat)
    print()
    if len(seen) >= 4:
        break
