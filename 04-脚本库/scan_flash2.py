# -*- coding: utf-8 -*-
"""扫描 Aspen 原生示例，取 Flash2 的规格写法（duty / vfrac）与吸收塔规格"""
import sys, os, re, glob
sys.stdout.reconfigure(encoding='utf-8')

roots = [r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples',
         r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Templates',
         r'C:\Program Files\AspenTech\Aspen Plus V15.0\Favorites']
files = []
for r in roots:
    for ext in ('*.bkp', '*.apwz'):
        files += glob.glob(os.path.join(r, '**', ext), recursive=True)
print('候选文件 %d' % len(files))

L = []
hit = 0
for f in files:
    try:
        if f.lower().endswith('.apwz'):
            import zipfile
            z = zipfile.ZipFile(f)
            data = b''
            for n in z.namelist():
                d = z.read(n)
                if b'BLOCK FLASH2' in d or b'BLOCK RADFRAC' in d:
                    data = d
                    break
            if not data:
                continue
            t = data.decode('latin-1')
        else:
            t = open(f, encoding='latin-1', errors='ignore').read()
    except Exception:
        continue
    if 'BLOCK FLASH2' not in t:
        continue
    hit += 1
    if hit > 8:
        break
    L.append('=' * 25 + ' ' + os.path.basename(f))
    starts = [s.start() for s in re.finditer(r'\?\s*BLOCK\s+FLASH2\s+', t)]
    for k in range(min(3, len(starts))):
        s0 = starts[k]
        s1 = starts[k + 1] if k + 1 < len(starts) else min(len(t), s0 + 900)
        L.append(re.sub(r'[ \t]+', ' ', t[s0:min(s1, s0 + 700)]))
        L.append('---')
    # 也找 FLOWSHEET 里的 FLASH2 连接
    for m in re.finditer(r'BLOCK BLKID = "[^"]+" BLKTYPE = "FLASH2"[^\\]{0,140}', t):
        L.append('FS: ' + re.sub(r'\s+', ' ', m.group()))
    L.append('')

open(r'D:\<化工工作区>\_probe\flash2_fmt.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('DONE 命中 %d' % hit)
