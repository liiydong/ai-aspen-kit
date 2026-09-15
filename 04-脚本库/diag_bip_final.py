# -*- coding: utf-8 -*-
"""综合诊断：NRTL 全字段 + UNIFAC 基团对照 + COM 设 ESTIMATE + 运行抓面板"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>'
REF = r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\Batch Modeling\Batch Distillation\WaterMethanol.bkp'

NRT = r'\Data\Properties\Parameters\Binary Interaction\NRTL-1'
UFG = r'\Data\Components\UNIFAC-Groups'


def list_children(doc, path, out, limit=80):
    nd = doc.Tree.FindNode(path)
    if nd is None:
        out.append('  %s -> None' % path)
        return
    try:
        c = nd.Elements.Count
    except Exception:
        out.append('  %s -> leaf' % path)
        return
    out.append('  %s  [%d]' % (path, c))
    for i in range(min(c, limit)):
        try:
            e = nd.Elements.Item(i)
        except Exception:
            break
        if e is None:
            continue
        try:
            nm = e.Name
        except Exception:
            nm = 'item%d' % i
        # 该子节点的元素数（行数）
        try:
            cc = e.Elements.Count
        except Exception:
            cc = '-'
        try:
            v = e.Value
        except Exception:
            v = '?'
        out.append('      %-16s rows=%-4s val=%r' % (nm, cc, v))


def open_doc(path, wait=30):
    msgs = []

    class Sink:
        def OnControlPanelMessage(self, *a):
            s = ' '.join(str(x) for x in a).strip()
            if s and s != 'False':
                msgs.append(s)

    doc = win32.DispatchEx('Apwn.Document')
    doc.SuppressDialogs = True
    win32.WithEvents(doc, Sink)
    doc.InitFromArchive2(path)
    t = time.time()
    while time.time() - t < wait:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.25)
    return doc, msgs


def run2(doc, msgs, secs=300):
    doc.Engine.Run2(False)
    t = time.time()
    while time.time() - t < secs:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.25)
        if time.time() - t > 10:
            try:
                if not bool(doc.Engine.IsRunning):
                    break
            except Exception:
                break
    for _ in range(80):
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)


# ============ 1) 参考模型 WaterMethanol ============
print('#' * 76)
print('# 参考模型 WaterMethanol（NRTL）UNIFAC 基团是否有值')
print('#' * 76)
doc, msgs = open_doc(REF, 25)
out = []
list_children(doc, UFG + r'\Input\GROUPNO', out, 30)
list_children(doc, NRT + r'\Input', out, 80)
print('\n'.join(out))
try:
    doc.Close()
except Exception:
    pass

# ============ 2) 我们的模型 ============
print()
print('#' * 76)
print('# NA_work：NRTL 全字段 + 设 ESTIMATE=YES + 运行')
print('#' * 76)
doc, msgs = open_doc(os.path.join(BASE, 'NA_work.bkp'), 30)
print('打开后消息数:', len(msgs))

out = []
list_children(doc, NRT + r'\Input', out, 80)
list_children(doc, UFG + r'\Input', out, 20)
print('\n'.join(out))

# 尝试设置 ESTIMATE = YES
print()
print('--- 尝试 COM 设 ESTIMATE=YES ---')
nd = doc.Tree.FindNode(NRT + r'\Input\ESTIMATE')
ok = False
if nd is not None:
    try:
        e0 = nd.Elements.Item(0)
        print('  当前:', e0.Name, '=', e0.Value)
        e0.Value = 'YES'
        ok = True
    except Exception as ex:
        print('  元素赋值失败:', str(ex)[:90])
    if not ok:
        try:
            nd.Value = 'YES'
            ok = True
        except Exception as ex:
            print('  节点赋值失败:', str(ex)[:90])
if ok:
    try:
        print('  读回:', nd.Elements.Item(0).Value)
    except Exception as ex:
        print('  读回失败:', str(ex)[:80])

# 运行
print()
print('--- 运行 ---')
n0 = len(msgs)
run2(doc, msgs, 300)
new = msgs[n0:]
print('运行期间新消息:', len(new))
for m in new:
    if re.search(r'ESTIMAT|UNIFAC|BINARY|PARAM|MISSING|GROUP|ERROR|TERMINAL|SEVERE|Warn', m, re.I):
        print('   |', m[:140])
print('--- 运行尾部 ---')
for m in new[-14:]:
    print('   >', m[:120])

# 运行后重读
out2 = []
list_children(doc, NRT + r'\Output', out2, 30)
list_children(doc, UFG + r'\Input\GROUPNO', out2, 30)
print()
print('\n'.join(out2))

try:
    doc.Close()
except Exception:
    pass
print('DONE')
