# -*- coding: utf-8 -*-
"""dump 示例文件的组分顺序 + NRTL-1 段原始格式"""
import sys, os, re, glob
sys.stdout.reconfigure(encoding='utf-8')

ROOTS = [r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples',
         r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Templates',
         r'C:\Program Files\AspenTech\Aspen Plus V15.0\Engine\Examples']


def find(name):
    for r in ROOTS:
        c = glob.glob(os.path.join(r, '**', name), recursive=True)
        if c:
            return c[0]
    return None


for name in ['teg.bkp', 'ICPEbegin.bkp', 'Conceptual Model Example.bkp', 'BenzeneToluene.bkp']:
    f = find(name)
    print('=' * 74)
    print(name, '->', f)
    if not f:
        continue
    s = open(f, encoding='latin-1', errors='ignore').read()
    # 组分顺序
    i = s.find('? COMPONENTS MAIN ?')
    j = s.find('/  CID =', i)
    seg = s[i:i+9000]
    cids = re.findall(r'CID = "?([A-Z0-9-]{1,20})"?\s+ANAME', seg)
    print('  组分(%d):' % len(cids), cids)
    # NRTL-1 段
    k = s.find('"NRTL-1"')
    print('  NRTL-1 @', k)
    if k >= 0:
        m = s.find('? PROPERTIES', k)
        blk = s[k:k+3000]
        print('  ---- 原文 ----')
        print('  ' + blk.replace('\n', '\n  ')[:2600])
    print()
