# -*- coding: utf-8 -*-
"""测试A：把 NRTL 参数对象的 ESTIMATE 打开，看能否自动补齐二元参数"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_Submit.bkp'
OUT = r'D:\<化工工作区>\_probe\est_yes.bkp'
t = open(SRC, encoding='utf-8', errors='ignore').read()
L = []

# 只打开 NRTL-1 段内的 ESTIMATE
m = re.search(r'PARAMETERS BINARY\s*\n?"T-DEPENDENT" "NRTL-1" \?', t)
L.append('NRTL-1 段定位: %s' % ('OK' if m else 'FAIL'))
if m:
    seg_start = m.start()
    nxt = re.search(r'\n\?\s*[A-Z]', t[m.end():])
    seg_end = m.end() + (nxt.start() if nxt else 2000)
    seg = t[seg_start:seg_end]
    L.append('原段内 ESTIMATE 出现: %s' % re.findall(r'ESTIMATE\s*=\s*\w+', seg))
    seg2 = re.sub(r'ESTIMATE\s*=\s*NO', 'ESTIMATE = YES', seg)
    t = t[:seg_start] + seg2 + t[seg_end:]
    L.append('改后: %s' % re.findall(r'ESTIMATE\s*=\s*\w+', seg2))
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)

MSGS = []
class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            MSGS.append(s)

doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
win32.WithEvents(doc, Sink)
doc.InitFromArchive2(OUT)
time.sleep(3)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 600:
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

L.append('')
L.append('===== 面板中含物理性质/参数/估算 的消息 =====')
for i, x in enumerate(MSGS):
    if re.search(r'PARAM|BINARY|ESTIMAT|UNIFAC|PROPERT|DATABANK|Terminal|Severe|Errors|Warn', x, re.I):
        L.append('  %3d | %s' % (i, x[:170]))

try:
    doc.SaveAs(r'D:\<化工工作区>\_probe\est_yes_s.bkp')
    doc.Close()
except Exception as ex:
    L.append('保存失败 %s' % ex)

# 统计保存后文件里的 NRTL 对
p2 = r'D:\<化工工作区>\_probe\est_yes_s.bkp'
if os.path.exists(p2):
    t2 = re.sub(r'\s+', ' ', open(p2, encoding='utf-8', errors='ignore').read())
    pairs = re.findall(r'PARAMNAME2 = NRTL CID1 = "?([A-Za-z0-9-]+)"? CID2 = "?([A-Za-z0-9-]+)"?', t2)
    L.append('')
    L.append('===== 保存后文件中的 NRTL 二元对: %d 条 =====' % len(pairs))
    for a, b in pairs:
        L.append('   %s - %s' % (a, b))
    for mm in re.finditer(r'BPVAL PARAMNAME2 = NRTL CID1 = "?([A-Za-z0-9-]+)"? CID2 = "?([A-Za-z0-9-]+)"?(.{0,500})', t2):
        src = sorted(set(re.findall(r'VAL\d+ = "([^"]+)"', mm.group(3))))
        L.append('   来源 %s - %s : %s' % (mm.group(1), mm.group(2), src))

open(r'D:\<化工工作区>\_probe\est_test.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('\n'.join(L))
