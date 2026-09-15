import re, sys
sys.stdout.reconfigure(encoding='utf-8')
for tag, p in [('模拟', r'D:/<化工工作区>/NA-Chemical-10000t_模拟.bkp'),
               ('sim2', r'D:/<化工工作区>/NA-Chemical-10000t_sim2.bkp'),
               ('sim3', r'D:/<化工工作区>/NA-Chemical-10000t_sim3.bkp'),
               ('工作版', r'D:/<化工工作区>/NA-Chemical-10000t_工作版.bkp')]:
    try:
        t = open(p, encoding='utf-8', errors='ignore').read()
    except Exception as e:
        print(tag, 'ERR', e); continue
    print('=' * 25, tag, 'chars =', len(t), 'lines =', t.count('\n'))
    secs = re.findall(r'\?\s*([A-Z][A-Z0-9_\- \\"./]{0,40}?)\s*\?', t)
    print('   段数:', len(secs))
    seen = []
    for s in secs:
        s = re.sub(r'\s+', ' ', s).strip()
        if s not in seen:
            seen.append(s)
    print('   前 30 个段:', seen[:30])
    print('   有 COMPONENTS:', '? COMPONENTS' in t, '| 有 BLOCK RSTOIC:', 'BLOCK RSTOIC' in t,
          '| 有 FLOWSHEET:', 'FLOWSHEET' in t)
    print()
