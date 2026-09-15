import re, sys, os
sys.stdout.reconfigure(encoding='utf-8')
for tag, p in [('sim3_saved', r'D:/<化工工作区>/_probe/sim3_saved.bkp'),
               ('sim2_saved', r'D:/<化工工作区>/_probe/sim2_saved.bkp')]:
    if not os.path.exists(p):
        print(tag, '不存在'); continue
    t = open(p, encoding='utf-8', errors='ignore').read()
    print('#' * 40, tag, len(t))
    for pat, name in [(r'\?\s*BLOCK\s+RSTOIC\s+"?R-101\s*"?\s*\?', 'R-101'),
                      (r'\?\s*BLOCK\s+RSTOIC\s+"?F-501\s*"?\s*\?', 'F-501'),
                      (r'\?\s*BLOCK\s+RADFRAC\s+"?T-401\s*"?\s*\?', 'T-401')]:
        m = re.search(pat, t)
        if not m:
            print('  %s 未找到' % name); continue
        nxt = re.search(r'\n\?\s*BLOCK\s', t[m.start() + 10:])
        end = m.start() + 10 + nxt.start() if nxt else min(len(t), m.start() + 3000)
        print('  ===== %s' % name)
        print(re.sub(r'\n{2,}', '\n', t[m.start():end]))
        print()
