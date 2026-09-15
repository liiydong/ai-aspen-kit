# -*- coding: utf-8 -*-
"""验证 NA_struct.bkp：完整性 + 结构是否可用 + 估算测试"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>'
P = os.path.join(BASE, 'NA_struct.bkp')

msgs = []
class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            msgs.append(s)

doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
win32.WithEvents(doc, Sink)
doc.InitFromArchive2(P)
t = time.time()
while time.time() - t < 30:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)

out = []
nd = doc.Tree.FindNode(r'\Data')
out.append('\\Data 子节点 = %s' % (nd.Elements.Count if nd else None))
for p in [r'\Data\Streams', r'\Data\Blocks']:
    n = doc.Tree.FindNode(p)
    out.append('  %s = %s' % (p, n.Elements.Count if n is not None else None))

out.append('--- 分子结构 ---')
for c in ['H2O', 'TOL', '3-MP', '3-CP', 'NAM']:
    n = doc.Tree.FindNode(r'\Data\Properties\Molecular Structure\%s\Input' % c)
    if n is None:
        out.append('  %s -> None' % c); continue
    rows = []
    try:
        for i in range(n.Elements.Count):
            e = n.Elements.Item(i)
            try:
                cc = e.Elements.Count
                if cc > 0:
                    rows.append('%s(%d)' % (e.Name, cc))
            except Exception:
                pass
    except Exception:
        pass
    out.append('  %-6s %s' % (c, ', '.join(rows[:6]) if rows else '(空)'))

print('\n'.join(out))
try:
    doc.Close()
except Exception:
    pass
print('DONE')
