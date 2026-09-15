# -*- coding: utf-8 -*-
"""v2：先写组成再写温压；并把全部成功/失败明细打出来"""
import time
import win32com.client as win32

MID = r'D:\<化工工作区>\_probe\setup1.bkp'
FIN = r'D:\<化工工作区>\NA-Chemical-10000t_模拟.bkp'

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(MID)
time.sleep(4)
ok, bad = [], []


def setv(path, val, tag):
    try:
        n = doc.Tree.FindNode(path)
        if n is None:
            bad.append('%s(无节点)' % tag)
            return
        n.Value = val
        ok.append(tag)
    except Exception as e:
        bad.append('%s(%s)' % (tag, str(e)[:38].replace('\n', ' ')))


setv(r'\Data\Properties\Specifications\Input\GOPSETNAME', 'NRTL', '物性方法')

PRES = {'S-101': (25, 2.0), 'S-102': (25, 2.0), 'S-103': (25, 2.0),
        'S-107': (25, 1.0), 'S-111': (40, 1.0), 'S-123': (25, 1.0),
        'S-124': (25, 1.0), 'S-125': (25, 1.0)}
FLOW = {'S-101': {'3-MP': 13.829, '4-MP': 0.070}, 'S-102': {'NH3': 13.906},
        'S-103': {'O2': 39.31, 'N2': 147.9}, 'S-107': {'H2O': 125.8},
        'S-111': {'TOL': 24.47}, 'S-123': {'NAOH': 0.065},
        'S-124': {'H2O': 159.5}, 'S-125': {'H2SO4': 0.0101}}

# ① 先组成
for s, comp in FLOW.items():
    for c, f in comp.items():
        setv(r'\Data\Streams' + '\\' + s + r'\Input\FLOW\MIXED' + '\\' + c, f,
             '%s.%s' % (s, c))
# ② 再温压
for s, (T, P) in PRES.items():
    setv(r'\Data\Streams' + '\\' + s + r'\Input\TEMP', T, '%s.T' % s)
    setv(r'\Data\Streams' + '\\' + s + r'\Input\PRES', P, '%s.P' % s)

# ③ 模块
BLK = {'M-101': [('PRES', 2.0)],
       'E-101': [('TEMP', 150.0), ('PRES', 1.8)],
       'E-104': [('TEMP', 100.0)],
       'E-201': [('TEMP', 40.0)],
       'E-601': [('TEMP', 50.0), ('PRES', 0.5)]}
for b, items in BLK.items():
    for f, v in items:
        setv(r'\Data\Blocks' + '\\' + b + r'\Input\\' + f, v, '%s.%s' % (b, f))

TWR = {'T-201': (6, 1.0, 'NONE', 'NONE'), 'T-301': (8, 1.0, None, None),
       'T-401': (25, 1.0, 'TOTAL', 'KETTLE'), 'T-402': (40, 1.0, 'TOTAL', 'KETTLE'),
       'T-403': (45, 0.60, 'TOTAL', 'KETTLE'), 'T-404': (32, 0.30, 'TOTAL', 'KETTLE')}
for b, (n, p, cond, reb) in TWR.items():
    setv(r'\Data\Blocks' + '\\' + b + r'\Input\NSTAGE', n, '%s.N' % b)
    setv(r'\Data\Blocks' + '\\' + b + r'\Input\PRES1', p, '%s.P1' % b)
    if cond:
        setv(r'\Data\Blocks' + '\\' + b + r'\Input\CONDENSER', cond, '%s.COND' % b)
        setv(r'\Data\Blocks' + '\\' + b + r'\Input\REBOILER', reb, '%s.REB' % b)

print('=== 成功 %d 项 ===' % len(ok), flush=True)
print(' ', ok, flush=True)
print('=== 失败 %d 项 ===' % len(bad), flush=True)
for x in bad:
    print('  ', x, flush=True)

# 验证
print()
for s in ['S-101', 'S-103']:
    for f in ['TEMP', 'PRES']:
        n = doc.Tree.FindNode(r'\Data\Streams' + '\\' + s + r'\Input\\' + f)
        print('  %s.%s = %r' % (s, f, n.Value if n is not None else None), flush=True)
for b in ['T-401', 'T-201']:
    for f in ['NSTAGE', 'PRES1', 'CONDENSER', 'REBOILER']:
        n = doc.Tree.FindNode(r'\Data\Blocks' + '\\' + b + r'\Input\\' + f)
        print('  %s.%s = %r' % (b, f, n.Value if n is not None else None), flush=True)

try:
    doc.SaveAs(FIN)
    print('已保存:', FIN, flush=True)
except Exception as e:
    print('SaveAs 失败:', str(e)[:90], flush=True)
try:
    doc.Close()
except Exception:
    pass
