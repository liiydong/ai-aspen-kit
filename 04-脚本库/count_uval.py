# -*- coding: utf-8 -*-
import glob, os, re

def nrtl_stats(p):
    t = open(p, encoding='utf-8', errors='ignore').read()
    i = t.find('PARAMNAME = NRTL')
    if i < 0:
        return None
    j = t.find('PARAMNAME =', i + 20)
    seg = t[i:j] if j > i else t[i:i+20000]
    return dict(size=len(t), bpval=seg.count('BPVAL'), uval=seg.count('UVAL'),
                cid=len(re.findall(r'CID1\s*=', seg)))

print('%-28s %9s %6s %5s %4s' % ('file', 'size', 'BPVAL', 'UVAL', 'CID1'))
for p in sorted(glob.glob(r'D:\<化工工作区>\_probe\db_*_s.bkp')) + \
         sorted(glob.glob(r'D:\<化工工作区>\_probe\nv_*.bkp')) + \
         [r'D:\<化工工作区>\NA_struct.bkp',
          r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Datapkg\glycols.bkp']:
    if not os.path.exists(p):
        continue
    s = nrtl_stats(p)
    if s is None:
        print('%-28s  (无 NRTL 段)' % os.path.basename(p))
    else:
        print('%-28s %9d %6d %5d %4d' % (os.path.basename(p), s['size'], s['bpval'], s['uval'], s['cid']))
