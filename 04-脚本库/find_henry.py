# -*- coding: utf-8 -*-
import glob, os, re

# 1) 亨利组分段的写法（从 Aspen 自带示例里找）
FILES = glob.glob(r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\**\*.bkp', recursive=True)
print('示例文件数:', len(FILES))
hit = 0
for p in FILES:
    try:
        t = open(p, encoding='utf-8', errors='ignore').read()
    except Exception:
        continue
    m = re.search(r'\?\s*COMPONENTS\s+HENRY-COMPS\s*\?', t)
    if not m:
        m = re.search(r'HENRY-COMPS', t)
    if m:
        seg = re.sub(r'\s+', ' ', t[max(0, m.start() - 60): m.start() + 420])
        print()
        print('==', os.path.basename(p))
        print('  ', seg)
        hit += 1
        if hit >= 4:
            break
if not hit:
    print('未在示例中找到 HENRY-COMPS 段')

# 2) 在我们自己的文件里找"检查结果/相稳定性检查"相关字段
print()
print('=== 我们文件 SETUP 段附近的候选字段 ===')
t = open(r'D:\<化工工作区>\NA-Chemical-10000t_BIP2.bkp', encoding='utf-8', errors='ignore').read()
for k in ['CHECK', 'CHK', 'PSTAB', 'STABCHK', 'PHCHK', 'RESULT', 'CALC-OPT']:
    ms = re.findall(r'[A-Z-]*%s[A-Z-]*' % k, t[:60000])
    u = sorted(set(ms))
    if u:
        print('  %-10s %s' % (k, u[:12]))

# 3) SETUP 段落原文
i = t.find('? SETUP')
print()
print('=== SETUP 段（前 1200 字符）===')
print(re.sub(r'[ \t]+', ' ', t[i:i+1200]))
