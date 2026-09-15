import re, sys
sys.stdout.reconfigure(encoding='utf-8')
t = open(r'D:/<化工工作区>/NA-Chemical-10000t_sim5.bkp', encoding='utf-8', errors='ignore').read()
m = re.search(r'\?\s*FLOWSHEET\s+[^?]*\?', t)
nxt = re.search(r'\n\?\s*[A-Z]', t[m.end():])
end = m.end() + nxt.start() if nxt else len(t)
seg = t[m.start():end]
flat = re.sub(r'\s+', ' ', seg)
for rec in re.split(r'\\ \\', flat):
    rec = rec.strip()
    if rec.startswith('\\'):
        rec = rec[1:].strip()
    if rec:
        print(rec)
