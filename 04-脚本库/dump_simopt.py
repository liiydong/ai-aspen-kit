# -*- coding: utf-8 -*-
import glob, os, re

FILES = glob.glob(r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\**\*.bkp', recursive=True)
shown = 0
for p in FILES:
    try:
        t = open(p, encoding='utf-8', errors='ignore').read()
    except Exception:
        continue
    m = re.search(r'\?\s*SETUP\s+"SIM-OPTIONS"\s*\?', t)
    if not m:
        continue
    seg = t[m.start(): m.start() + 700]
    print('==', os.path.basename(p))
    print(re.sub(r'[ \t]+', ' ', seg))
    print()
    shown += 1
    if shown >= 3:
        break

# 全库搜含 FREE-WATER / NPHASE / PHASE-STAB 的 SIM-OPTIONS 原文
print('=== 全库搜 SETUP 段内的 FREE-WATER / NPHASE ===')
hit = 0
for p in FILES:
    try:
        t = open(p, encoding='utf-8', errors='ignore').read()
    except Exception:
        continue
    for k in ['FREE-WATER =', 'NPHASE =', 'PHASE-STAB']:
        for mm in re.finditer(re.escape(k) + r'\s*\S+', t):
            ctx = re.sub(r'\s+', ' ', t[max(0, mm.start()-90): mm.start()+60])
            if 'APPLICABLE' in ctx:
                continue
            print('  %-40s %s' % (os.path.basename(p)[:38], ctx))
            hit += 1
            break
        if hit >= 8:
            break
    if hit >= 8:
        break
if not hit:
    print('  未找到')
