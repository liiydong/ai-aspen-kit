# -*- coding: utf-8 -*-
"""最终：构建可运行模型 -> 运行 -> 导出报表与结果 -> 保存带结果的模型"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_sim6.bkp'
WORK = r'D:\<化工工作区>\NA-Chemical-10000t_final.bkp'
REP = r'D:\<化工工作区>\_probe\final_report.txt'
SUM = r'D:\<化工工作区>\_probe\final_summary.txt'
MSG = r'D:\<化工工作区>\_probe\final_panel.txt'

t = open(BASE, encoding='utf-8', errors='ignore').read()
t = t.replace('"COL-CONFIG" CONDENSER = NONE REBOILER = NONE',
              '"COL-CONFIG" CONDENSER = NONE REBOILER = KETTLE')
m = re.search(r'"COL-SPECS" BASIS-RDV = 1\.0 <0> <0> ', t)
t = t[:m.start()] + '"COL-SPECS" BASIS-RDV = 1.0 <0> <0> BASIS-BR = 0.1 <-1> <0> ' + t[m.end():]
m2 = re.search(r'PARAM\s+BASE\s*=\s*[^\s\\]+', t)
t = t[:m2.start()] + 'PARAM BASE = "PENG-ROB"' + t[m2.end():]
open(WORK, 'w', encoding='utf-8', errors='ignore').write(t)
print('工作模型已生成:', WORK, flush=True)

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
doc.InitFromArchive2(WORK)
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
print('=== 控制面板尾部 ===', flush=True)
for x in MSGS[-22:]:
    print('   |', x[:190], flush=True)
open(MSG, 'w', encoding='utf-8').write('\n'.join(MSGS))


def g(p):
    n = doc.Tree.FindNode(p)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return 'ERR'


print()
print('=== 模块结果 ===', flush=True)
lines = []
for b in ['M-101', 'E-101', 'R-101', 'E-104', 'T-201', 'E-201', 'T-301', 'T-401',
          'T-402', 'T-403', 'T-404', 'R-501', 'F-501', 'E-601']:
    bb = r'\Data\Blocks\%s\Output' % b
    d = {f: g(bb + '\\' + f) for f in
         ['TEMP', 'PRES', 'DUTY', 'COND_DUTY', 'REB_DUTY', 'MOLE_RR', 'B-PRES',
          'BOTTOM_TEMP', 'TOP_TEMP', 'MOLEFLOW']}
    s = '  %-7s ' % b + ' '.join('%s=%s' % (k, v) for k, v in d.items() if v not in (None, '', 0))
    print(s, flush=True)
    lines.append(s)

print()
print('=== 导出报表 ===', flush=True)
for typ, path, tag in [(6, MSG, 'messages'), (3, SUM, 'summary'), (2, REP, 'report')]:
    try:
        doc.Engine.Export(typ, path)
        sz = os.path.getsize(path) if os.path.exists(path) else -1
        print('  Export(%d) -> %s  (%d bytes)' % (typ, tag, sz), flush=True)
    except Exception as ex:
        print('  Export(%d) 失败: %s' % (typ, ex), flush=True)

print()
print('=== 保存 ===', flush=True)
for p in [r'D:\<化工工作区>\NA-Chemical-10000t_final.bkp',
          r'D:\<化工工作区>\NA-Chemical-10000t_final.apwz']:
    try:
        doc.SaveAs(p)
        print('  已保存:', p, os.path.getsize(p), flush=True)
    except Exception as ex:
        print('  保存失败 %s: %s' % (p, ex), flush=True)

try:
    doc.Close()
except Exception:
    pass
print('DONE', flush=True)
