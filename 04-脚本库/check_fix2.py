# -*- coding: utf-8 -*-
"""跑 fix2 看还剩哪些错 + 找 EXTRACT 示例"""
import os, re, time, glob
import pythoncom
import win32com.client as win32

# --- 1) 跑 fix2，抓控制面板 ---
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
except Exception as e:
    print('事件钩子失败', str(e)[:80], flush=True)

doc.InitFromArchive2(r'D:\<化工工作区>\_probe\fix2.bkp')
time.sleep(4)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 200:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 10:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(40):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)
print('=== fix2 运行 %.0fs，控制面板 %d 条 ===' % (time.time() - t0, len(MSGS)), flush=True)
for m in MSGS:
    print('   ', m[:300], flush=True)
print()
print('=== 运行结果抽查 ===', flush=True)
for b in ['M-101', 'E-101', 'R-101', 'T-201', 'T-301', 'T-401', 'R-501', 'F-501']:
    n = doc.Tree.FindNode(r'\Data\Blocks\%s\Output' % b)
    if n is None:
        print('  %-7s 无' % b, flush=True)
        continue
    vals = []
    for f in ['TEMP', 'PRES', 'DUTY', 'MOLE_RR']:
        try:
            x = n.Elements.Item(f)
            if x is not None and x.Value not in (None, ''):
                vals.append('%s=%s' % (f, x.Value))
        except Exception:
            pass
    print('  %-7s %s' % (b, ','.join(vals) if vals else '(空)'), flush=True)
try:
    doc.SaveAs(r'D:\<化工工作区>\_probe\fix2_saved.bkp')
except Exception:
    pass
try:
    doc.Close()
except Exception:
    pass

# --- 2) 找 EXTRACT 示例段 ---
print()
print('=== 搜索 Aspen 示例里的 EXTRACT 段 ===', flush=True)
root = r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples'
n_found = 0
for p in glob.glob(os.path.join(root, '**', '*.bkp'), recursive=True):
    try:
        t = open(p, encoding='utf-8', errors='ignore').read()
    except Exception:
        continue
    if 'BLOCK EXTRACT' not in t:
        continue
    m = re.search(r'\?\s*BLOCK\s+EXTRACT\s+\S+\s*\?', t)
    if not m:
        continue
    m2 = re.search(r'\?\s+[A-Z]', t[m.end():])
    end = m.end() + (m2.start() if m2 else 1200)
    print('---', os.path.basename(p), flush=True)
    print(re.sub(r'\s+', ' ', t[m.start():end])[:1500], flush=True)
    n_found += 1
    if n_found >= 3:
        break
if n_found == 0:
    print('  示例中未找到 EXTRACT 段', flush=True)
