# -*- coding: utf-8 -*-
"""用用户回存的带结构文件：检查完整性 + 试估算二元参数"""
import sys, re, time, os, shutil
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>'
SRC = r'C:\Users\<用户名>\Desktop\烟酰胺Aspen模型_请补参数后另存.bkp'
W = os.path.join(BASE, 'NA_user.bkp')
shutil.copy(SRC, W)
print('已复制 -> NA_user.bkp  %d bytes' % os.path.getsize(W))

msgs = []


class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            msgs.append(s)


doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
win32.WithEvents(doc, Sink)
doc.InitFromArchive2(W)
t = time.time()
while time.time() - t < 30:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)

print()
print('=== 打开后完整性检查 ===')
nd = doc.Tree.FindNode(r'\Data')
print('  \\Data 子节点数 =', nd.Elements.Count if nd else None)
for p in [r'\Data\Streams', r'\Data\Blocks', r'\Data\Flowsheet']:
    n = doc.Tree.FindNode(p)
    print('  %-20s = %s' % (p, n.Elements.Count if n is not None else None))

print()
print('=== 分子结构状态 ===')
for c in ['H2O', 'TOL', '3-MP', '3-CP', 'NAM', 'CO']:
    n = doc.Tree.FindNode(r'\Data\Properties\Molecular Structure\%s\Input' % c)
    if n is None:
        print('  %-6s -> None' % c); continue
    info = []
    for col in ['ATOMNO1', 'DISPATOM1', 'BONDTYPE1', 'FORMULA']:
        nn = doc.Tree.FindNode(r'\Data\Properties\Molecular Structure\%s\Input\%s' % (c, col))
        if nn is not None:
            try:
                info.append('%s=%d' % (col, nn.Elements.Count))
            except Exception:
                pass
    # 扫一遍所有子节点，找有行的
    rows = []
    try:
        for i in range(n.Elements.Count):
            e = n.Elements.Item(i)
            try:
                nm, c2 = e.Name, e.Elements.Count
                if c2 > 0:
                    rows.append('%s(%d)' % (nm, c2))
            except Exception:
                pass
    except Exception:
        pass
    print('  %-6s 有数据的列: %s' % (c, ', '.join(rows[:8]) if rows else '(全空)'))

print()
print('=== 二元参数状态 ===')
for p in [r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Input',
          r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Output\CID1',
          r'\Data\Properties\Parameters\Binary Interaction\UNIQ-1\Output\CID1',
          r'\Data\Components\UNIFAC-Groups\Input\GROUPNO',
          r'\Data\Properties\Parameters\UNIFAC Groups']:
    n = doc.Tree.FindNode(p)
    if n is None:
        print('  %-64s None' % p.split('Binary Interaction')[-1]); continue
    try:
        print('  %-64s rows=%d' % (p.split('Binary Interaction')[-1], n.Elements.Count))
    except Exception:
        print('  %-64s ?' % p)

# 设估算开关
print()
print('=== 设 ESTIMATE=YES + ALLONLY=ALL ===')
for p, v in [(r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Input\ESTIMATE', 'YES'),
             (r'\Data\Properties\Estimation\Estimate\Input\ALLONLY', 'ALL')]:
    n = doc.Tree.FindNode(p)
    try:
        n.Elements.Item(0).Value = v
        print('  %s -> %r' % (p.split('\\')[-1], n.Elements.Item(0).Value))
    except Exception as e:
        print('  %s 失败: %s' % (p.split('\\')[-1], str(e)[:70]))

n0 = len(msgs)
print()
print('=== Run ===')
doc.Engine.Run2(False)
t = time.time()
while time.time() - t < 360:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.3)
    if time.time() - t > 10:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(100):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)
new = msgs[n0:]
print('  新消息 %d 条' % len(new))
est = [m for m in new if re.search(r'ESTIMAT|UNIFAC|PCES|MISSING', m)]
print('  估算相关 %d 条:' % len(est))
for m in est[:10]:
    print('     |', m[:125])
for m in new[-8:]:
    print('     >', m[:110])

print()
print('=== 运行后参数状态 ===')
for p in [r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Input\VAL1',
          r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Output\CID1',
          r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Output\VALUE',
          r'\Data\Properties\Parameters\Binary Interaction\UNIQ-1\Output\CID1',
          r'\Data\Components\UNIFAC-Groups\Input\GROUPNO']:
    n = doc.Tree.FindNode(p)
    if n is None:
        print('  %-58s None' % p.split('Binary Interaction')[-1]); continue
    try:
        print('  %-58s rows=%d' % (p.split('Binary Interaction')[-1], n.Elements.Count))
    except Exception:
        print('  %-58s ?' % p)

out = os.path.join(BASE, 'NA_user_est.bkp')
try:
    doc.SaveAs(out)
    print()
    print('SaveAs ->', os.path.basename(out), os.path.getsize(out))
except Exception as e:
    print('SaveAs 失败:', str(e)[:80])
try:
    doc.Close()
except Exception:
    pass
print('DONE')
