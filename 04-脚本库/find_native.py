import glob, os, re, sys
sys.stdout.reconfigure(encoding='utf-8')

roots = [r'C:/Program Files/AspenTech/Aspen Plus V15.0',
         r'C:/Program Files/Common Files/AspenTech Shared',
         r'C:/ProgramData/AspenTech']
found = []
for r in roots:
    for ext in ('bkp', 'apw', 'inp'):
        for p in glob.glob(os.path.join(r, '**', '*.' + ext), recursive=True):
            try:
                with open(p, 'r', encoding='utf-8', errors='ignore') as f:
                    head = f.read(400000)
                if 'STOIC REACNO' in head:
                    found.append(p)
            except Exception:
                pass
print('含 STOIC REACNO 的文件数:', len(found))
for p in found[:6]:
    print('  ', p)

# 挑一个原生文件，打印原始 repr
if found:
    p = found[0]
    t = open(p, encoding='utf-8', errors='ignore').read()
    m = re.search(r'\?\s*BLOCK\s+RSTOIC', t)
    print()
    print('#' * 60)
    print('# 原生文件:', os.path.basename(p))
    print('# RAW REPR (2000 chars)')
    print(repr(t[m.start():m.start() + 2000]))
