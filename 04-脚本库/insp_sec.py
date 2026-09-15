import re, io, sys
sys.stdout.reconfigure(encoding='utf-8')

P = r'D:/<化工工作区>/NA-Chemical-10000t_working.bkp'
import os
if not os.path.exists(P):
    P = r'D:/<化工工作区>/NA-Chemical-10000t_模拟.bkp'
print('FILE =', P)
t = open(P, encoding='utf-8', errors='ignore').read()
print('len =', len(t))

def show(pat, name, limit=1600):
    print('=' * 30, name)
    for m in re.finditer(pat, t):
        j = t.find('? ', m.end())
        # find next section marker
        k = m.end()
        nxt = re.search(r'\n\?\s', t[k:])
        end = k + nxt.start() if nxt else k + limit
        seg = t[m.start():min(end, m.start() + limit)]
        seg = re.sub(r'[ \t]+', ' ', seg)
        print('--- @%d ---' % m.start())
        print(seg)
        print()

show(r'\?\s*BLOCK\s+RSTOIC\s+"?R-101', 'R-101 RSTOIC')
show(r'\?\s*BLOCK\s+RADFRAC\s+"?T-201', 'T-201 RADFRAC')
show(r'\?\s*BLOCK\s+EXTRACT\s+"?T-301', 'T-301 EXTRACT')
print('==== 全部 BLOCK 段标记 ====')
for m in re.finditer(r'\?\s*BLOCK\s+[A-Z0-9]+\s+"?[A-Z0-9-]+"?', t):
    print('  %7d  %s' % (m.start(), re.sub(r'\s+', ' ', m.group())))
