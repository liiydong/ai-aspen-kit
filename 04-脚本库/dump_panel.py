# -*- coding: utf-8 -*-
"""导出模型全部控制面板消息（逐条），供人工判读警告全文"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_BIP2.bkp'
OUTF = r'D:\<化工工作区>\_probe\panel_all.txt'
MSGS = []


class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            MSGS.append(s)


doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
win32.WithEvents(doc, Sink)
doc.InitFromArchive2(SRC)
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
for _ in range(90):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)

with open(OUTF, 'w', encoding='utf-8') as f:
    f.write('total %d\n' % len(MSGS))
    for i, m in enumerate(MSGS):
        f.write('%4d | %s\n' % (i, m))

# 单独抽出警告块（WARNING + 后 4 行）
blocks = []
for i, m in enumerate(MSGS):
    if re.match(r'^\s*\*\s*WARNING', m) or re.match(r'^\s*\*\s*ERROR', m):
        blk = [MSGS[j] for j in range(i, min(i + 5, len(MSGS)))]
        blocks.append((i, blk))
with open(r'D:\<化工工作区>\_probe\warn_blocks.txt', 'w', encoding='utf-8') as f:
    f.write('warn blocks: %d\n\n' % len(blocks))
    for i, blk in blocks:
        f.write('--- @%d ---\n' % i)
        for b in blk:
            f.write('   %s\n' % b[:170])
        f.write('\n')

try:
    doc.Close()
except Exception:
    pass
try:
    doc.Quit()
except Exception:
    pass
print('messages:', len(MSGS), 'blocks:', len(blocks))
