import glob, os, re, sys
sys.stdout.reconfigure(encoding='utf-8')

roots = [r'C:/Program Files/AspenTech/Aspen Plus V15.0',
         r'C:/Program Files/Common Files/AspenTech Shared',
         r'C:/ProgramData/AspenTech',
         r'D:/<化工工作区>/_probe']
nat_rad, nat_ext = [], []
for r in roots:
    for p in glob.glob(os.path.join(r, '**', '*.bkp'), recursive=True):
        try:
            t = open(p, encoding='utf-8', errors='ignore').read(1500000)
        except Exception:
            continue
        if 'BLOCK RADFRAC' in t and len(nat_rad) < 3:
            nat_rad.append(p)
        if 'BLOCK EXTRACT' in t and len(nat_ext) < 3:
            nat_ext.append(p)

print('RADFRAC 候选:', nat_rad)
print('EXTRACT 候选:', nat_ext)
print()


def dump(p, kind, name):
    t = open(p, encoding='utf-8', errors='ignore').read()
    m = re.search(r'\?\s*BLOCK\s+%s\s+\S+' % kind, t)
    if not m:
        print('  未找到', name); return
    nxt = re.search(r'\n\?\s*BLOCK\s', t[m.start() + 10:])
    end = m.start() + 10 + nxt.start() if nxt else min(len(t), m.start() + 2500)
    seg = t[m.start():min(end, m.start() + 2500)]
    print('#' * 55)
    print('#', name, '|', os.path.basename(p))
    print(repr(seg))
    print()


for p in nat_rad[:2]:
    dump(p, 'RADFRAC', '原生 RADFRAC')
for p in nat_ext[:2]:
    dump(p, 'EXTRACT', '原生 EXTRACT')

# 多效精馏 参考
ref = r'D:/<化工工作区>/_probe/ref_multi.bkp'
if os.path.exists(ref):
    dump(ref, 'RADFRAC', '多效精馏 RADFRAC')
