# -*- coding: utf-8 -*-
import glob, os, re

FILES = glob.glob(r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\**\*.bkp', recursive=True)
seen = {}
for p in FILES:
    try:
        t = open(p, encoding='utf-8', errors='ignore').read()
    except Exception:
        continue
    m = re.search(r'ODATABANKS[^\\]{0,300}', t)
    if m:
        s = re.sub(r'\s+', ' ', m.group(0))
        seen.setdefault(s, []).append(os.path.basename(p))

print('不同的 ODATABANKS 写法:', len(seen))
for s, fl in sorted(seen.items(), key=lambda x: -len(x[1])):
    print()
    print('[%d 个文件] %s' % (len(fl), s[:280]))
    print('   例:', fl[:4])

# 找 LOADDECH
print()
print('=== 含 LOADDECH 的文件 ===')
for p in FILES:
    try:
        t = open(p, encoding='utf-8', errors='ignore').read()
    except Exception:
        continue
    if 'LOADDECH' in t:
        i = t.find('LOADDECH')
        print(' ', os.path.basename(p), '::', re.sub(r'\s+', ' ', t[max(0,i-120):i+120]))
        break
else:
    print('  无')
