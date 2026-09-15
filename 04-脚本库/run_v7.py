# -*- coding: utf-8 -*-
"""v7：补齐 R-101 副反应（4-MP 氨氧化 / 深度氧化 / HCN / CO）+ T-201 再沸器 + PENG-ROB
   并把流股结果正确读出（TEMP_OUT / MOLEFLOW / MASSFLMX / 组分流量）"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_sim6.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_v7.bkp'
NL = '\n'

R101 = NL.join([
    '? BLOCK RSTOIC "R-101" ? ; "METCBAR_MOLE" ; ; ICON1 ; ',
    '\\ PARAM TEMP = 405.0 <22> <4> PRES = 1.8 <20> <5> SPEC-OPT = TP ',
    # 反应1 主反应 3-MP 氨氧化
    '\\ \\ STOIC REACNO = 1 STOIC-CID = "3-MP" STOIC-SSID = MIXED COEF = -1.0 <0> <0> ',
    '\\ \\ STOIC REACNO = 1 STOIC-CID = "NH3" STOIC-SSID = MIXED COEF = -1.0 <0> <0> ',
    '\\ \\ STOIC REACNO = 1 STOIC-CID = "O2" STOIC-SSID = MIXED COEF = -1.5 <0> <0> ',
    # 反应2 4-MP 氨氧化（新增！）
    '\\ \\ STOIC REACNO = 2 STOIC-CID = "4-MP" STOIC-SSID = MIXED COEF = -1.0 <0> <0> ',
    '\\ \\ STOIC REACNO = 2 STOIC-CID = "NH3" STOIC-SSID = MIXED COEF = -1.0 <0> <0> ',
    '\\ \\ STOIC REACNO = 2 STOIC-CID = "O2" STOIC-SSID = MIXED COEF = -1.5 <0> <0> ',
    # 反应3 深度氧化
    '\\ \\ STOIC REACNO = 3 STOIC-CID = "3-MP" STOIC-SSID = MIXED COEF = -1.0 <0> <0> ',
    '\\ \\ STOIC REACNO = 3 STOIC-CID = "O2" STOIC-SSID = MIXED COEF = -7.75 <0> <0> ',
    # 反应4 HCN 生成
    '\\ \\ STOIC REACNO = 4 STOIC-CID = "3-MP" STOIC-SSID = MIXED COEF = -1.0 <0> <0> ',
    '\\ \\ STOIC REACNO = 4 STOIC-CID = "O2" STOIC-SSID = MIXED COEF = -6.5 <0> <0> ',
    # 反应5 CO 生成（脱羰）
    '\\ \\ STOIC REACNO = 5 STOIC-CID = "3-MP" STOIC-SSID = MIXED COEF = -1.0 <0> <0> ',
    '\\ \\ STOIC REACNO = 5 STOIC-CID = "O2" STOIC-SSID = MIXED COEF = -4.75 <0> <0> ',
    # 产物
    '\\ \\ STOIC1 REACNO1 = 1 STOIC-CID1 = "3-CP" STOIC-SSID1 = MIXED COEF1 = 1.0 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 1 STOIC-CID1 = "H2O" STOIC-SSID1 = MIXED COEF1 = 3.0 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 2 STOIC-CID1 = "4-CP" STOIC-SSID1 = MIXED COEF1 = 1.0 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 2 STOIC-CID1 = "H2O" STOIC-SSID1 = MIXED COEF1 = 3.0 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 3 STOIC-CID1 = "CO2" STOIC-SSID1 = MIXED COEF1 = 6.0 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 3 STOIC-CID1 = "N2" STOIC-SSID1 = MIXED COEF1 = 0.5 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 3 STOIC-CID1 = "H2O" STOIC-SSID1 = MIXED COEF1 = 3.5 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 4 STOIC-CID1 = "CO2" STOIC-SSID1 = MIXED COEF1 = 5.0 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 4 STOIC-CID1 = "H2O" STOIC-SSID1 = MIXED COEF1 = 3.0 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 4 STOIC-CID1 = "HCN" STOIC-SSID1 = MIXED COEF1 = 1.0 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 5 STOIC-CID1 = "CO" STOIC-SSID1 = MIXED COEF1 = 6.0 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 5 STOIC-CID1 = "N2" STOIC-SSID1 = MIXED COEF1 = 0.5 <0> <0> ',
    '\\ \\ STOIC1 REACNO1 = 5 STOIC-CID1 = "H2O" STOIC-SSID1 = MIXED COEF1 = 3.5 <0> <0> ',
    # 转化率
    '\\ \\ CONVEX EXT-REACNO = 1 KEY-SSID = MIXED KEY-CID = "3-MP" CONV = .88 <0> <0> ',
    '\\ \\ CONVEX EXT-REACNO = 2 KEY-SSID = MIXED KEY-CID = "4-MP" CONV = .88 <0> <0> ',
    '\\ \\ CONVEX EXT-REACNO = 3 KEY-SSID = MIXED KEY-CID = "3-MP" CONV = .07 <0> <0> ',
    '\\ \\ CONVEX EXT-REACNO = 4 KEY-SSID = MIXED KEY-CID = "3-MP" CONV = .016 <0> <0> ',
    '\\ \\ CONVEX EXT-REACNO = 5 KEY-SSID = MIXED KEY-CID = "3-MP" CONV = .006 <0> <0> ',
    '\\ \\ PRODUCTS SID = "S-106" \\ ',
])

t = open(BASE, encoding='utf-8', errors='ignore').read()
m = re.search(r'\?\s*BLOCK\s+RSTOIC\s+"?R-101\s*"?\s*\?', t)
nxt = re.search(r'\n\?\s*BLOCK\s', t[m.start() + 10:])
end = m.start() + 10 + nxt.start() if nxt else len(t)
t = t[:m.start()] + R101 + NL + t[end:]

t = t.replace('"COL-CONFIG" CONDENSER = NONE REBOILER = NONE',
              '"COL-CONFIG" CONDENSER = NONE REBOILER = KETTLE')
mm = re.search(r'"COL-SPECS" BASIS-RDV = 1\.0 <0> <0> ', t)
t = t[:mm.start()] + '"COL-SPECS" BASIS-RDV = 1.0 <0> <0> BASIS-BR = 0.1 <-1> <0> ' + t[mm.end():]
m3 = re.search(r'PARAM\s+BASE\s*=\s*[^\s\\]+', t)
t = t[:m3.start()] + 'PARAM BASE = "PENG-ROB"' + t[m3.end():]
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)
print('已生成:', OUT, flush=True)

MSGS = []


class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            MSGS.append(s)


doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
try:
    win32.WithEvents(doc, Sink)
except Exception:
    pass
doc.InitFromArchive2(OUT)
time.sleep(3)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 480:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 10:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(80):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)

print()
print('=== 控制面板（错误/警告相关） ===', flush=True)
for x in MSGS:
    if any(k in x for k in ['ERROR', 'SEVERE', 'WARNING', 'stopped', 'Converged',
                            'MASS BALANCE', 'completed']):
        print('   |', x[:190], flush=True)


def g(p):
    n = doc.Tree.FindNode(p)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return 'ERR'


print()
print('=== 流股汇总（TEMP_OUT/PRES_OUT/MOLEFLOW/MASSFLMX） ===', flush=True)
for s in ['S-104', 'S-106', 'S-108', 'S-109', 'S-110', 'S-112', 'S-113', 'S-114',
          'S-115', 'S-116', 'S-117', 'S-118', 'S-119', 'S-120', 'S-121', 'S-122',
          'S-201', 'S-301', 'S-401', 'S-402']:
    b = r'\Data\Streams\%s\Output' % s
    print('  %-7s T=%s P=%s N=%s W=%s BETA=%s' % (
        s, g(b + r'\TEMP_OUT'), g(b + r'\PRES_OUT'), g(b + r'\MOLEFLOW'),
        g(b + r'\MASSFLMX'), g(b + r'\BETA')), flush=True)

print()
print('=== 找组分流量节点（S-106 Output 里元素数>0 的子节点） ===', flush=True)
n = doc.Tree.FindNode(r'\Data\Streams\S-106\Output')
if n is not None:
    for i in range(n.Elements.Count):
        try:
            e = n.Elements.Item(i)
            try:
                c = e.Elements.Count
            except Exception:
                c = -1
            if c > 0:
                print('   子节点 %s [%d]' % (e.Name, c), flush=True)
                for j in range(min(c, 20)):
                    try:
                        e2 = e.Elements.Item(j)
                        v2 = None
                        try:
                            v2 = e2.Value
                        except Exception:
                            pass
                        print('        .%s = %s' % (e2.Name, v2), flush=True)
                    except Exception:
                        pass
        except Exception:
            pass

print()
print('=== 模块结果 ===', flush=True)
for b in ['T-201', 'T-301', 'T-401', 'T-402', 'T-403', 'T-404']:
    bb = r'\Data\Blocks\%s\Output' % b
    print('  %-7s COND=%s REB=%s RR=%s Ttop=%s Tbot=%s' % (
        b, g(bb + r'\COND_DUTY'), g(bb + r'\REB_DUTY'), g(bb + r'\MOLE_RR'),
        g(bb + r'\TOP_TEMP'), g(bb + r'\BOTTOM_TEMP')), flush=True)

for p in [r'D:\<化工工作区>\NA-Chemical-10000t_v7.bkp', r'D:\<化工工作区>\NA-Chemical-10000t_v7.apwz']:
    try:
        doc.SaveAs(p)
        print('  已保存:', p, os.path.getsize(p), flush=True)
    except Exception as ex:
        print('  保存失败:', p, ex, flush=True)
try:
    doc.Close()
except Exception:
    pass
print('DONE', flush=True)
