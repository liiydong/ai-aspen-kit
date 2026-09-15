import re, sys
sys.stdout.reconfigure(encoding='utf-8')
for tag, p in [('我们的 conv', r'D:/<化工工作区>/NA-Chemical-10000t_conv.bkp'),
               ('原生 多效精馏', r'D:/<化工工作区>/_probe/ref_multi.bkp')]:
    t = open(p, encoding='utf-8', errors='ignore').read()
    print('=' * 30, tag)
    m = re.search(r'\?\s*PROPERTIES\s+"OPTION-SETS"', t)
    if not m:
        print('未找到'); continue
    nxt = re.search(r'\n\?\s*PROPERTIES\s+"MOLEC|\\ \?\s*PROPERTIES', t[m.end():])
    end = m.end() + (nxt.start() if nxt else 1500)
    print(re.sub(r'[ \t]+', ' ', t[m.start():min(end, m.start() + 1600)]))
    print()
    print('文件里 BASE = 的写法:')
    for mm in re.finditer(r'PARAM\s*\n?\s*BASE\s*=\s*[^\n]{0,40}', t):
        print('   ', repr(mm.group(0)))
