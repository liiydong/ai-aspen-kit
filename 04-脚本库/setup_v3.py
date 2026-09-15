# -*- coding: utf-8 -*-
"""v3：补 OPTION-SETS 定义物性方法；修正路径；重设全部参数"""
import time, re
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\setup1.bkp'      # 已含三个反应段
MID = r'D:\<化工工作区>\_probe\setup3.bkp'
FIN = r'D:\<化工工作区>\NA-Chemical-10000t_模拟.bkp'

t = open(SRC, encoding='utf-8', errors='ignore').read()

old = 'GPROPERTIES PARCON = -1'
new = ('GPROPERTIES GPPROCTYPE = ALL GBASEOPSET = "NRTL" GOPSETNAME = "NRTL" '
       'PARCON = -2 \\ ? PROPERTIES "OPTION-SETS" "NRTL" ? \\ PARAM BASE = "NRTL" \\')
# 注意原文本里 GPROPERTIES 与 PARCON 之间可能跨行
m = re.search(r'GPROPERTIES\s+PARCON = -1', t)
if not m:
    print('未找到 GPROPERTIES 锚点！', flush=True)
    print(repr(t[t.find('? PROPERTIES MAIN ?'):t.find('? PROPERTIES MAIN ?') + 200]), flush=True)
    raise SystemExit(1)
t = t[:m.start()] + new + t[m.end():]
open(MID, 'w', encoding='utf-8', errors='ignore').write(t)
print('OPTION-SETS 已注入', flush=True)
i = t.find('? PROPERTIES MAIN ?')
print('  片段:', repr(re.sub(r'\s+', ' ', t[i:i + 260])), flush=True)

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(MID)
time.sleep(4)

ok, bad = [], []


def S(path, val, tag):
    try:
        n = doc.Tree.FindNode(path)
        if n is None:
            bad.append('%s(无节点)' % tag)
            return
        n.Value = val
        got = doc.Tree.FindNode(path).Value
        if got is None:
            bad.append('%s(写后仍为None)' % tag)
        else:
            ok.append('%s=%s' % (tag, got))
    except Exception as e:
        bad.append('%s(%s)' % (tag, str(e)[:45].replace('\n', ' ')))


FLOW = {'S-101': {'3-MP': 13.829, '4-MP': 0.070}, 'S-102': {'NH3': 13.906},
        'S-103': {'O2': 39.31, 'N2': 147.9}, 'S-107': {'H2O': 125.8},
        'S-111': {'TOL': 24.47}, 'S-123': {'NAOH': 0.065},
        'S-124': {'H2O': 159.5}, 'S-125': {'H2SO4': 0.0101}}
TP = {'S-101': (25, 2.0), 'S-102': (25, 2.0), 'S-103': (25, 2.0),
      'S-107': (25, 1.0), 'S-111': (40, 1.0), 'S-123': (25, 1.0),
      'S-124': (25, 1.0), 'S-125': (25, 1.0)}

for s, comp in FLOW.items():
    for c, f in comp.items():
        S(r'\Data\Streams\%s\Input\FLOW\MIXED\%s' % (s, c), f, '%s.%s' % (s, c))
# 测一个温度看 AE_UNDERSPEC 是否消失
S(r'\Data\Streams\S-101\Input\TEMP', 25.0, 'S-101.T')
S(r'\Data\Streams\S-101\Input\PRES', 2.0, 'S-101.P')
print()
print('=== 成功 %d ===' % len(ok), flush=True)
print(' ', ok, flush=True)
print('=== 失败 %d ===' % len(bad), flush=True)
for x in bad:
    print('  ', x, flush=True)

BLK = {'M-101': [('PRES', 2.0)], 'E-101': [('TEMP', 150.0), ('PRES', 1.8)],
       'E-104': [('TEMP', 100.0)], 'E-201': [('TEMP', 40.0)],
       'E-601': [('TEMP', 50.0), ('PRES', 0.5)]}
print()
for b, items in BLK.items():
    for f, v in items:
        S(r'\Data\Blocks\%s\Input\%s' % (b, f), v, '%s.%s' % (b, f))
TWR = {'T-201': (6, 1.0), 'T-301': (8, 1.0), 'T-401': (25, 1.0),
       'T-402': (40, 1.0), 'T-403': (45, 0.60), 'T-404': (32, 0.30)}
for b, (n, p) in TWR.items():
    S(r'\Data\Blocks\%s\Input\NSTAGE' % b, n, '%s.N' % b)
    S(r'\Data\Blocks\%s\Input\PRES1' % b, p, '%s.P1' % b)
print('模块成功 %d:' % len(ok), ok[-24:], flush=True)
print('模块失败 %d:' % len(bad), bad[-20:], flush=True)
try:
    doc.SaveAs(FIN)
    print('已保存', FIN, flush=True)
except Exception as e:
    print('SaveAs 失败', str(e)[:90], flush=True)
try:
    doc.Close()
except Exception:
    pass
