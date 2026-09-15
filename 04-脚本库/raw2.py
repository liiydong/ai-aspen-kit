import re, sys, os
sys.stdout.reconfigure(encoding='utf-8')

def dump(tag, path, key, n=1500):
    print('#' * 60)
    print('#', tag, '|', os.path.basename(path))
    if not os.path.exists(path):
        print('  文件不存在'); return
    t = open(path, encoding='utf-8', errors='ignore').read()
    print('  size =', len(t))
    m = re.search(key, t)
    if not m:
        print('  未匹配:', key); return
    print('  at', m.start())
    print('--- RAW REPR ---')
    print(repr(t[m.start():m.start() + n]))
    print()

# 1. Aspen 原生多行 STOIC（cumene.bkp）
dump('原生 cumene RSTOIC', r'C:/Program Files/AspenTech/Aspen Plus V15.0/GUI/Examples/cumene.bkp',
     r'\?\s*BLOCK\s+RSTOIC\s+\w+', 1200)

# 2. 我们的 fix3.bkp 里的 R-101 / F-501 / T-201 / T-301
F = r'D:/<化工工作区>/_probe/fix3.bkp'
dump('我们 R-101', F, r'\?\s*BLOCK\s+RSTOIC\s+"?R-101', 1000)
dump('我们 F-501', F, r'\?\s*BLOCK\s+RSTOIC\s+"?F-501', 900)
dump('我们 T-201', F, r'\?\s*BLOCK\s+RADFRAC\s+"?T-201', 1100)
dump('我们 T-301', F, r'\?\s*BLOCK\s+EXTRACT\s+"?T-301', 1100)
