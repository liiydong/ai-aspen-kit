import glob, os, re, sys
sys.stdout.reconfigure(encoding='utf-8')

roots = [r'C:/Program Files/AspenTech/Aspen Plus V15.0', r'C:/ProgramData/AspenTech']
hits = 0
for r in roots:
    for p in glob.glob(os.path.join(r, '**', '*.bkp'), recursive=True):
        try:
            t = open(p, encoding='utf-8', errors='ignore').read(3000000)
        except Exception:
            continue
        # 找块级物性方法写法
        for m in re.finditer(r'.{0,90}OPSETNAME\s*=.{0,60}', t):
            s = re.sub(r'\s+', ' ', m.group(0))
            if 'GBASEOPSET' in s and 'GOPSETNAME' not in s:
                continue
            print('%-40s %s' % (os.path.basename(p)[:38], s))
            hits += 1
            break
        if hits > 12:
            break
    if hits > 12:
        break
print('---- 命中', hits)

print()
print('==== 我们文件里 OPSETNAME 的用法 ====')
t = open(r'D:/<化工工作区>/NA-Chemical-10000t_v7.bkp', encoding='utf-8', errors='ignore').read()
for m in re.finditer(r'.{0,80}OPSETNAME\s*=.{0,50}', t):
    print('  ', re.sub(r'\s+', ' ', m.group(0)))

print()
print('==== 找 PROP-SET / "OPSET" 相关关键字 ====')
for kw in ['"BLOCK-OPTION"', '"BLOCK-OPTIONS"', 'GOPSETNAME', 'PROP-OPTION', '"PR-SEC"']:
    print('  %-16s %d' % (kw, t.count(kw)))
