# -*- coding: utf-8 -*-
"""dump 各模块 Input 的字段名，找 STOIC/CONVEX 等反应字段"""
import time
import win32com.client as win32

P = r'D:\<化工工作区>\NA-Chemical-10000t_工作版.bkp'
OUT = r'D:\<化工工作区>\_probe\block_fields.txt'

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(P)
time.sleep(4)

lines = []
for b in ['R-101', 'R-501', 'F-501', 'T-201', 'T-301', 'T-401', 'T-402', 'T-403',
          'T-404', 'M-101', 'E-101', 'E-104', 'E-201', 'E-601']:
    n = doc.Tree.FindNode(r'\Data\Blocks' + '\\' + b + r'\Input')
    if n is None:
        lines.append('%s : None' % b)
        continue
    lines.append('=' * 12 + ' %s  [%d 子]' % (b, n.Elements.Count))
    for i in range(n.Elements.Count):
        ch = n.Elements.Item(i)
        try:
            c = ch.Elements.Count
        except Exception:
            c = '-'
        try:
            v = repr(ch.Value)[:40]
        except Exception:
            v = ''
        lines.append('   %-32s 子=%-5s %s' % (ch.Name, c, v))
    lines.append('')

open(OUT, 'w', encoding='utf-8').write('\n'.join(lines))
print('已写', OUT, len(lines), '行')

txt = '\n'.join(lines)
print()
print('=== 含关键字的模块 ===')
for kw in ['STOIC', 'CONVEX', 'REAC', 'RXN', 'PARAM', 'TEMP', 'PRES', 'NSTAGE',
           'FEED', 'COND', 'REB', 'RR', 'RATE', 'SPEC', 'FLASH', 'NPHASE', 'DUTY']:
    hits = [l.strip() for l in lines if kw in l and l.startswith('   ')]
    if hits:
        print('  [%s] %d 个:' % (kw, len(hits)))
        for h in hits[:8]:
            print('        ', h)
try:
    doc.Close()
except Exception:
    pass
