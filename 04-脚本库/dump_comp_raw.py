# -*- coding: utf-8 -*-
import re

P = r'D:\<化工工作区>\NA_struct.bkp'
t = open(P, encoding='utf-8', errors='ignore').read()
i = t.find('? COMPONENTS MAIN ?')
j = t.find('PROPERTIES', i)
seg = t[i:j]
# 去掉折行，便于观察每条记录
flat = re.sub(r'\s+', ' ', seg)
recs = [r.strip() for r in flat.split('/') if r.strip()]
print('records:', len(recs))
for r in recs:
    cid = re.search(r'CID\s*=\s*("[^"]*"|\S+)', r)
    typ = re.search(r'TYPE\s*=\s*(\S+)', r)
    db1 = re.search(r'DBNAME1\s*=\s*("[^"]*"|\S+)', r)
    print('%-10s TYPE=%-8s DBNAME1=%-28s | %s' % (
        cid.group(1).strip('"') if cid else '?',
        typ.group(1) if typ else 'MISSING',
        db1.group(1).strip('"') if db1 else '-',
        r[:150]))
