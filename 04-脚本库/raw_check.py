import re, sys
sys.stdout.reconfigure(encoding='utf-8')

P = r'D:/<化工工作区>/NA-Chemical-10000t_模拟.bkp'
t = open(P, encoding='utf-8', errors='ignore').read()

for name, key in [('R-101', r'\?\s*BLOCK\s+RSTOIC\s+"?R-101'),
                  ('T-201', r'\?\s*BLOCK\s+RADFRAC\s+"?T-201'),
                  ('T-301', r'\?\s*BLOCK\s+EXTRACT\s+"?T-301')]:
    m = re.search(key, t)
    print('=' * 25, name, 'at', m.start())
    seg = t[m.start():m.start() + 700]
    # 显示真实换行
    print(repr(seg))
    print()

print('==== 统计：文件中真实换行的行数 ====')
print('  总行数:', t.count('\n'))
lines = t.split('\n')
print('  最长行长度:', max(len(l) for l in lines))
import collections
c = collections.Counter(len(l) for l in lines)
print('  行长度分布 top10:', c.most_common(10))
