import re, sys
sys.stdout.reconfigure(encoding='utf-8')
t = open(r'D:/<化工工作区>/NA-Chemical-10000t_模拟.bkp', encoding='utf-8', errors='ignore').read()
starts = [m.start() for m in re.finditer(r'\?\s*BLOCK\s+RADFRAC\s+"?(T-40[1-4])"?\s*\?', t)]
allb = [m.start() for m in re.finditer(r'\?\s*BLOCK\s+[A-Z0-9]+\s+"?[A-Za-z0-9-]+"?\s*\?', t)]
for s in starts:
    e = min([x for x in allb if x > s] or [len(t)])
    print('=' * 60)
    print(repr(t[s:e]))
    print()
