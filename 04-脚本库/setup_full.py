# -*- coding: utf-8 -*-
"""化学法 Aspen 完整建模：反应注入 + 流股 + 模块参数 + 物性方法 -> 保存运行"""
import time, sys
import win32com.client as win32

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_工作版.bkp'
MID = r'D:\<化工工作区>\_probe\setup1.bkp'
FIN = r'D:\<化工工作区>\NA-Chemical-10000t_模拟.bkp'

# ============ 1) 文本注入三个反应器 ============
t = open(SRC, encoding='utf-8', errors='ignore').read()

R101 = ('? BLOCK RSTOIC "R-101" ? ; "METCBAR_MOLE" ; ; ICON1 ; '
        '\\ PARAM TEMP = 405.0 PRES = 1.8 SPEC-OPT = TP '
        '\\ STOIC REACNO = 1 STOIC-CID = "3-MP" STOIC-SSID = MIXED COEF = -1.0 <0> <0> '
        '\\ STOIC1 REACNO1 = 1 STOIC-CID1 = NH3 STOIC-SSID1 = MIXED COEF1 = -1.0 <0> <0> '
        '\\ STOIC2 REACNO2 = 1 STOIC-CID2 = O2 STOIC-SSID2 = MIXED COEF2 = -1.5 <0> <0> '
        '\\ STOIC3 REACNO3 = 1 STOIC-CID3 = "3-CP" STOIC-SSID3 = MIXED COEF3 = 1.0 <0> <0> '
        '\\ STOIC4 REACNO4 = 1 STOIC-CID4 = H2O STOIC-SSID4 = MIXED COEF4 = 3.0 <0> <0> '
        '\\ STOIC5 REACNO5 = 2 STOIC-CID5 = "3-MP" STOIC-SSID5 = MIXED COEF5 = -1.0 <0> <0> '
        '\\ STOIC6 REACNO6 = 2 STOIC-CID6 = O2 STOIC-SSID6 = MIXED COEF6 = -7.5 <0> <0> '
        '\\ STOIC7 REACNO7 = 2 STOIC-CID7 = CO2 STOIC-SSID7 = MIXED COEF7 = 6.0 <0> <0> '
        '\\ STOIC8 REACNO8 = 2 STOIC-CID8 = H2O STOIC-SSID8 = MIXED COEF8 = 3.0 <0> <0> '
        '\\ STOIC9 REACNO9 = 2 STOIC-CID9 = HCN STOIC-SSID9 = MIXED COEF9 = 1.0 <0> <0> '
        '\\ CONVEX EXT-REACNO = 1 KEY-SSID = MIXED KEY-CID = "3-MP" CONV = .86 <0> <0> '
        '\\ CONVEX1 EXT-REACNO1 = 2 KEY-SSID1 = MIXED KEY-CID1 = "3-MP" CONV1 = .02 <0> <0> '
        '\\ PRODUCTS SID = "S-106" \\ ')

R501 = ('? BLOCK RSTOIC "R-501" ? ; "METCBAR_MOLE" ; ; ICON1 ; '
        '\\ PARAM TEMP = 140.0 PRES = 4.0 SPEC-OPT = TP '
        '\\ STOIC REACNO = 1 STOIC-CID = "3-CP" STOIC-SSID = MIXED COEF = -1.0 <0> <0> '
        '\\ STOIC1 REACNO1 = 1 STOIC-CID1 = H2O STOIC-SSID1 = MIXED COEF1 = -1.0 <0> <0> '
        '\\ STOIC2 REACNO2 = 1 STOIC-CID2 = NAM STOIC-SSID2 = MIXED COEF2 = 1.0 <0> <0> '
        '\\ STOIC3 REACNO3 = 2 STOIC-CID3 = NAM STOIC-SSID3 = MIXED COEF3 = -1.0 <0> <0> '
        '\\ STOIC4 REACNO4 = 2 STOIC-CID4 = H2O STOIC-SSID4 = MIXED COEF4 = -1.0 <0> <0> '
        '\\ STOIC5 REACNO5 = 2 STOIC-CID5 = NAC STOIC-SSID5 = MIXED COEF5 = 1.0 <0> <0> '
        '\\ STOIC6 REACNO6 = 2 STOIC-CID6 = NH3 STOIC-SSID6 = MIXED COEF6 = 1.0 <0> <0> '
        '\\ CONVEX EXT-REACNO = 1 KEY-SSID = MIXED KEY-CID = "3-CP" CONV = .985 <0> <0> '
        '\\ CONVEX1 EXT-REACNO1 = 2 KEY-SSID1 = MIXED KEY-CID1 = NAM CONV1 = .005 <0> <0> '
        '\\ PRODUCTS SID = "S-201" \\ ')

F501 = ('? BLOCK RSTOIC "F-501" ? ; "METCBAR_MOLE" ; ; ICON1 ; '
        '\\ PARAM TEMP = 60.0 PRES = 1.0 SPEC-OPT = TP '
        '\\ STOIC REACNO = 1 STOIC-CID = NAOH STOIC-SSID = MIXED COEF = -2.0 <0> <0> '
        '\\ STOIC1 REACNO1 = 1 STOIC-CID1 = H2SO4 STOIC-SSID1 = MIXED COEF1 = -1.0 <0> <0> '
        '\\ STOIC2 REACNO2 = 1 STOIC-CID2 = NA2SO4 STOIC-SSID2 = MIXED COEF2 = 1.0 <0> <0> '
        '\\ STOIC3 REACNO3 = 1 STOIC-CID3 = H2O STOIC-SSID3 = MIXED COEF3 = 2.0 <0> <0> '
        '\\ CONVEX EXT-REACNO = 1 KEY-SSID = MIXED KEY-CID = H2SO4 CONV = 1.0 <0> <0> '
        '\\ PRODUCTS SID = "S-301" \\ ')

a = t.find('? BLOCK RSTOIC "F-501" ?')
t = t[:a] + R101 + '\n' + R501 + '\n' + F501 + '\n' + t[a:]
open(MID, 'w', encoding='utf-8', errors='ignore').write(t)
print('反应段已注入:', MID, flush=True)

# ============ 2) COM 设置其余参数 ============
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
            bad.append('%s (节点不存在)' % tag)
            return
        n.Value = val
        ok.append(tag)
    except Exception as e:
        bad.append('%s (%s)' % (tag, str(e)[:50]))


# 物性方法
setv(r'\Data\Properties\Specifications\Input\GOPSETNAME', 'NRTL', '物性方法=NRTL')

# 常压/加压单位：bar。压力按 bar 绝对压填
PRES = {  # 流股：T(℃), P(bar)
    'S-101': (25, 2.0), 'S-102': (25, 2.0), 'S-103': (25, 2.0),
    'S-107': (25, 1.0), 'S-111': (40, 1.0),
    'S-123': (25, 1.0), 'S-124': (25, 1.0), 'S-125': (25, 1.0),
}
FLOW = {  # kmol/h，按 MIXED 摩尔流量
    'S-101': {'3-MP': 13.829, '4-MP': 0.070},
    'S-102': {'NH3': 13.906},
    'S-103': {'O2': 39.31, 'N2': 147.9},
    'S-107': {'H2O': 125.8},
    'S-111': {'TOL': 24.47},
    'S-123': {'NAOH': 0.065},
    'S-124': {'H2O': 159.5},
    'S-125': {'H2SO4': 0.0101},
}
for s, (T, P) in PRES.items():
    setv(r'\Data\Streams' + '\\' + s + r'\Input\TEMP', T, '%s TEMP' % s)
    setv(r'\Data\Streams' + '\\' + s + r'\Input\PRES', P, '%s PRES' % s)
    for c, f in FLOW[s].items():
        setv(r'\Data\Streams' + '\\' + s + r'\Input\FLOW\MIXED' + '\\' + c, f,
             '%s %s' % (s, c))

# 模块参数
BLK = {
    'M-101': [('PRES', 2.0)],
    'E-101': [('TEMP', 150.0), ('PRES', 1.8)],
    'E-104': [('TEMP', 100.0)],
    'E-201': [('TEMP', 40.0)],
    'E-601': [('TEMP', 50.0), ('PRES', 0.5)],
}
for b, items in BLK.items():
    for f, v in items:
        setv(r'\Data\Blocks' + '\\' + b + r'\Input\\' + f, v, '%s.%s' % (b, f))

# 塔：理论板数 + 塔顶压力 + 冷凝器/再沸器
TWR = {
    'T-201': (6, 1.0, 'NONE', 'NONE'),
    'T-301': (8, 1.0, None, None),
    'T-401': (25, 1.0, 'TOTAL', 'KETTLE'),
    'T-402': (40, 1.0, 'TOTAL', 'KETTLE'),
    'T-403': (45, 0.60, 'TOTAL', 'KETTLE'),
    'T-404': (32, 0.30, 'TOTAL', 'KETTLE'),
}
for b, (n, p, cond, reb) in TWR.items():
    setv(r'\Data\Blocks' + '\\' + b + r'\Input\NSTAGE', n, '%s.NSTAGE' % b)
    setv(r'\Data\Blocks' + '\\' + b + r'\Input\PRES1', p, '%s.PRES1' % b)
    if cond:
        setv(r'\Data\Blocks' + '\\' + b + r'\Input\CONDENSER', cond, '%s.COND' % b)
        setv(r'\Data\Blocks' + '\\' + b + r'\Input\REBOILER', reb, '%s.REB' % b)

print('成功 %d 项：' % len(ok), ok[:12], flush=True)
print('失败 %d 项：' % len(bad), bad[:15], flush=True)

try:
    doc.SaveAs(FIN)
    print('已保存:', FIN, flush=True)
except Exception as e:
    print('SaveAs 失败:', str(e)[:100], flush=True)
try:
    doc.Close()
except Exception:
    pass
