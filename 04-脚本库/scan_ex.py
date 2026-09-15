# -*- coding: utf-8 -*-
"""扫描 Aspen 自带示例：找带 NRTL 二元参数数据的文件"""
import sys, os, re, glob
sys.stdout.reconfigure(encoding='utf-8')

ROOTS = [r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples',
         r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Templates',
         r'C:\Program Files\AspenTech\Aspen Plus V15.0\Engine\Examples']

files = []
for r in ROOTS:
    for ext in ('*.bkp', '*.apw', '*.apwz'):
        files += glob.glob(os.path.join(r, '**', ext), recursive=True)
files = sorted(set(files))
print('示例文件数:', len(files))
print()

hits = []
for f in files:
    try:
        if f.lower().endswith('.apwz'):
            import zipfile
            z = zipfile.ZipFile(f)
            txt = b''
            for n in z.namelist():
                d = z.read(n)
                if b'NRTL' in d:
                    txt = d
                    break
            if not txt:
                continue
            s = txt.decode('latin-1', 'ignore')
        else:
            s = open(f, encoding='latin-1', errors='ignore').read()
    except Exception:
        continue
    if 'NRTL' not in s:
        continue
    # 找 NRTL-1 段并看是否含数据记录（含 PARAM 行或含数字对）
    i = s.find('"NRTL-1"')
    hasdata = False
    nrec = 0
    if i >= 0:
        j = s.find('? ', i + 12)
        blk = s[i:i+30000] if j < 0 else s[i:j]
        nrec = blk.count('PARAM')
        # 数据记录通常带 \ \ ALL 或含 DSET / 成对数字
        hasdata = bool(re.search(r'NRTL\s+[-\d]', blk)) or ('APV150' in blk and nrec > 0)
    # 组分名
    comps = sorted(set(re.findall(r'CID = "?([A-Z0-9-]{2,20})"?', s)))[:0]
    if i >= 0 and (hasdata or nrec > 0):
        hits.append((f, nrec, len(s)))

print('=== 含 NRTL-1 段的示例 ===')
for f, nrec, ln in sorted(hits, key=lambda x: -x[1])[:25]:
    print('  %5d PARAM  %8d bytes  %s' % (nrec, ln, os.path.basename(f)))
print()
print('总命中:', len(hits))
