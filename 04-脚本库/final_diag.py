# -*- coding: utf-8 -*-
"""最终诊断：运行 + 看输出流股节点结构"""
import time
import win32com.client as win32

F = r'D:\<化工工作区>\_probe\full2.bkp'
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(F)
time.sleep(4)

r = None
try:
    r = doc.Engine.Run2()
    print('Run2 返回:', repr(r), flush=True)
except Exception as e:
    print('Run2 异常:', str(e)[:150], flush=True)
time.sleep(2)

n = doc.Tree.FindNode(r'\Data\Streams\S-104')
print()
print('S-104 子节点:', [n.Elements.Item(i).Name for i in range(n.Elements.Count)], flush=True)
for sub in ['Input', 'Output']:
    nn = doc.Tree.FindNode(r'\Data\Streams\S-104\%s' % sub)
    if nn is None:
        print('  %s: None' % sub)
        continue
    names = [nn.Elements.Item(i).Name for i in range(nn.Elements.Count)]
    print('  %s 字段数=%d，含 TEMP/PRES 的:' % (sub, len(names)),
          [x for x in names if 'TEMP' in x.upper() or 'PRES' in x.upper()][:6], flush=True)

print()
print('=== 结果汇总节点 ===', flush=True)
rs = doc.Tree.FindNode(r'\Data\Results Summary')
if rs is not None:
    for i in range(rs.Elements.Count):
        ch = rs.Elements.Item(i)
        try:
            c = ch.Elements.Count
        except Exception:
            c = '-'
        print('   %-24s %s' % (ch.Name, c), flush=True)
print()
print('=== 收敛节点 ===', flush=True)
cv = doc.Tree.FindNode(r'\Data\Convergence')
if cv is not None:
    for i in range(cv.Elements.Count):
        ch = cv.Elements.Item(i)
        try:
            c = ch.Elements.Count
        except Exception:
            c = '-'
        print('   %-24s %s' % (ch.Name, c), flush=True)
try:
    doc.Close()
except Exception:
    pass
