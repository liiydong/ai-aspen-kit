import re, sys
sys.stdout.reconfigure(encoding='utf-8')
for tag, p in [('多效精馏', r'D:/<化工工作区>/_probe/ref_multi.bkp'),
               ('我们的', r'D:/<化工工作区>/NA-Chemical-10000t_sim6.bkp')]:
    t = open(p, encoding='utf-8', errors='ignore').read()
    print('=' * 30, tag)
    for m in re.finditer(r'BLOCK\s+BLKID\s*=\s*"?([A-Za-z0-9_-]+)"?\s+BLKTYPE\s*=\s*"([^"]+)"\s+'
                         r'MDLTYPE\s*=\s*"([^"]+)"\s+IN\s*=\s*\(([^)]*)\)\s+OUT\s*=\s*\(([^)]*)\)', t):
        bid, bt, mt, inn, out = m.groups()
        if mt in ('RadFrac', 'Extract', 'RStoic'):
            print('  %-8s %-9s IN=[%s]' % (bid, mt, re.sub(r'\s+', ' ', inn).strip()))
            print('  %-8s %-9s OUT=[%s]' % ('', '', re.sub(r'\s+', ' ', out).strip()))
