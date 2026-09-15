# -*- coding: utf-8 -*-
"""诊断 4lib 报错 + 检查 UNIFAC-GROUP 段 + 遍历二元参数节点（正确路径）"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>'
P4 = os.path.join(BASE, 'NA_fix_4lib.bkp')

# ---------- 文本部分 ----------
for p in [P4, os.path.join(BASE, 'NA_fixres_4lib.bkp')]:
    if not os.path.exists(p):
        print('缺失', p); continue
    x = open(p, encoding='utf-8', errors='ignore').read()
    print('=' * 70)
    print(os.path.basename(p), len(x))
    i = x.find('DATABANKS ?')
    print('DATABANKS:', repr(x[max(0, i-40): i+330]))
    k = x.find('UNIFAC-GROUP')
    print('UNIFAC-GROUP 段:', repr(x[k:k+260]) if k > 0 else '无')
    j = x.find('UFGRP')
    print('含 UFGRP:', j > 0, repr(x[j:j+200]) if j > 0 else '')

# ---------- COM 部分：抓 4lib 报错 + 遍历节点 ----------
print()
print('=' * 70)
print('COM: 打开 4lib 抓报错')
msgs = []


class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            msgs.append(s)


doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
win32.WithEvents(doc, Sink)
doc.InitFromArchive2(P4)
time.sleep(4)
for _ in range(120):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)

print('消息总数:', len(msgs))
for m in msgs[:60]:
    print('  |', m[:150])

# 遍历二元参数节点（正确路径）
print()
print('--- 二元参数节点遍历 ---')
for path in [r'\Data\Properties\Parameters\Binary Interaction',
             r'\Data\Properties\Parameters\Binary Interaction\NRTL-1',
             r'\Data\Properties\Parameters']:
    nd = doc.Tree.FindNode(path)
    print('%s -> %s' % (path, 'None' if nd is None else 'OK'))
    if nd is not None:
        try:
            c = nd.Elements.Count
            print('   子节点数:', c)
            for i in range(min(c, 25)):
                try:
                    print('     -', nd.Elements.Item(i).Name)
                except Exception:
                    pass
        except Exception as e:
            print('   ', str(e)[:80])

try:
    doc.Close()
except Exception:
    pass
print('DONE')
