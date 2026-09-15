import re, sys
sys.stdout.reconfigure(encoding='utf-8')
for tag, p in [('sim2', r'D:/<化工工作区>/NA-Chemical-10000t_sim2.bkp'),
               ('sim3', r'D:/<化工工作区>/NA-Chemical-10000t_sim3.bkp')]:
    t = open(p, encoding='utf-8', errors='ignore').read()
    print('#' * 30, tag, 'size', len(t))
    for bid in ['R-101', 'F-501']:
        m = re.search(r'\?\s*BLOCK\s+RSTOIC\s+"?%s\s*"?\s*\?' % bid, t)
        if not m:
            print('  %s 未找到' % bid); continue
        nxt = re.search(r'\n\?\s*BLOCK\s', t[m.start() + 10:])
        end = m.start() + 10 + nxt.start() if nxt else len(t)
        seg = t[m.start():end]
        print('  ---', bid, 'len', len(seg))
        for i, row in enumerate(re.split(r'\\\\', seg)):
            row = row.strip()
            if row:
                print('      %2d| %s' % (i, row[:170]))
    print()
