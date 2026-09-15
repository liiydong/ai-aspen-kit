# -*- coding: utf-8 -*-
import re, glob, os

CAND = [
    r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\3phase.bkp',
]
CAND += glob.glob(r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\**\*.bkp', recursive=True)[:200]

def scan(p):
    try:
        t = open(p, encoding='utf-8', errors='ignore').read()
    except Exception:
        return None
    if 'PARAMNAME = NRTL' not in t:
        return None
    i = t.find('PARAMNAME = NRTL')
    # 找该段结束
    j = t.find('PARAMNAME =', i + 20)
    seg = t[i:j] if j > i else t[i:i+3000]
    n_uval = seg.count('UVAL')
    n_bpval = seg.count('BPVAL')
    n_est = len(re.findall(r'ESTIMATE\s*=\s*(\w+)', seg))
    return dict(path=p, size=len(t), bpval=n_bpval, uval=n_uval, estimate=n_est)

found = []
for p in CAND:
    r = scan(p)
    if r:
        found.append(r)

found.sort(key=lambda x: -x['uval'])
print('含 NRTL 段的示例文件数:', len(found))
for r in found[:15]:
    print('  UVAL=%-4d BPVAL=%-4d %s' % (r['uval'], r['bpval'], os.path.basename(r['path'])))

# 详细看最优的一个
if found:
    best = found[0]['path']
    t = open(best, encoding='utf-8', errors='ignore').read()
    i = t.find('PARAMNAME = NRTL')
    j = t.find('PARAMNAME =', i + 20)
    print()
    print('=== %s 的 NRTL 段（2000 字符）===' % os.path.basename(best))
    print(re.sub(r'[ \t]+', ' ', t[i:i+2000]))
