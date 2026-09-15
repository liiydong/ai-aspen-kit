# -*- coding: utf-8 -*-
import glob, os, re

FILES = glob.glob(r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\**\*.bkp', recursive=True)
names = {}
for p in FILES:
    try:
        t = open(p, encoding='utf-8', errors='ignore').read()
    except Exception:
        continue
    for m in re.finditer(r'\?\s*SETUP\s+"?([A-Z0-9 \-]+)"?\s*\?', t):
        names.setdefault(m.group(1).strip(), set()).add(os.path.basename(p))

print('=== 全部 SETUP 段名 ===')
for k, v in sorted(names.items()):
    print('  %-24s %d 文件   例: %s' % (k, len(v), list(v)[:2]))

print()
print('=== 含 CHECK / PSTAB / STAB 关键字的段名 ===')
for p in FILES[:250]:
    try:
        t = open(p, encoding='utf-8', errors='ignore').read()
    except Exception:
        continue
    for m in re.finditer(r'\?\s*"([A-Z0-9 \-]*(?:CHECK|PSTAB|STAB)[A-Z0-9 \-]*)"\s*\?', t):
        print('  ', os.path.basename(p), '->', m.group(1))
