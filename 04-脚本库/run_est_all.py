# -*- coding: utf-8 -*-
"""设 ALLONLY 打开全局估算 + NRTL ESTIMATE=YES，Run 后看二元参数是否生成"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>'
NRT = r'\Data\Properties\Parameters\Binary Interaction\NRTL-1'
EST = r'\Data\Properties\Estimation\Estimate\Input'

msgs = []


class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            msgs.append(s)


doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
win32.WithEvents(doc, Sink)
doc.InitFromArchive2(os.path.join(BASE, 'NA_work.bkp'))
t = time.time()
while time.time() - t < 25:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)


def setval(path, val):
    nd = doc.Tree.FindNode(path)
    if nd is None:
        return 'node None'
    try:
        e = nd.Elements.Item(0)
        old = e.Value
        e.Value = val
        return 'ok: %r -> %r' % (old, e.Value)
    except Exception as ex:
        try:
            old = nd.Value
            nd.Value = val
            return 'ok2: %r -> %r' % (old, nd.Value)
        except Exception as ex2:
            return 'FAIL %s | %s' % (str(ex)[:50], str(ex2)[:50])


print('--- 设置估算开关 ---')
for p, v in [(EST + r'\ALLONLY', 'ALL'),
             (EST + r'\BINCHOICE', 'YES'),
             (EST + r'\GRPCHOICE', 'YES'),
             (NRT + r'\Input\ESTIMATE', 'YES')]:
    print('  %-58s -> %s' % (p.split('\\')[-1], setval(p, v)))

# 读回
print()
print('--- 读回 ---')
for p in [EST + r'\ALLONLY', EST + r'\BINCHOICE', NRT + r'\Input\ESTIMATE']:
    nd = doc.Tree.FindNode(p)
    try:
        print('  %-20s = %r' % (p.split('\\')[-1], nd.Elements.Item(0).Value))
    except Exception as ex:
        print('  %-20s ?' % p.split('\\')[-1])

# 运行
print()
print('--- Run ---')
n0 = len(msgs)
doc.Engine.Run2(False)
t = time.time()
while time.time() - t < 400:
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
print('新消息:', len(new))
for m in new:
    if re.search(r'ESTIMAT|UNIFAC|BINARY|PCES|MISSING|GROUP|Terminal|Severe|Errors|Warnings', m):
        print('   |', m[:135])
print('--- 尾部 ---')
for m in new[-12:]:
    print('   >', m[:120])

# 读结果
print()
print('--- 二元参数节点 ---')
for p in [NRT + r'\Input\BDBANK', NRT + r'\Input\ESTIMATE',
          NRT + r'\Output\CID1', NRT + r'\Output\CID2', NRT + r'\Output\VALUE',
          r'\Data\Properties\Parameters\UNIFAC Groups',
          r'\Data\Components\UNIFAC-Groups\Input\GROUPNO']:
    nd = doc.Tree.FindNode(p)
    if nd is None:
        print('  %-56s None' % p); continue
    try:
        c = nd.Elements.Count
        print('  %-56s rows=%d' % (p, c))
    except Exception:
        try:
            print('  %-56s val=%r' % (p, nd.Value))
        except Exception:
            print('  %-56s ?' % p)

# 保存
out = os.path.join(BASE, 'NA_est_all.bkp')
try:
    doc.SaveAs(out)
    print('SaveAs ->', os.path.basename(out), os.path.getsize(out))
except Exception as e:
    print('SaveAs 失败:', str(e)[:80])
try:
    doc.Close()
except Exception:
    pass
print('DONE')
