# -*- coding: utf-8 -*-
"""决定性实验：参考模型改回流比后运行，看结果是否重算"""
import time
import win32com.client as win32

REF = r'D:\<化工工作区>\_probe\ref_multi.bkp'


def read_out(doc, tag):
    vals = {}
    for p in ['\\Data\\Blocks\\B1\\Output\\COND_DUTY',
              '\\Data\\Blocks\\B1\\Output\\REB_DUTY',
              '\\Data\\Blocks\\B1\\Output\\MOLE_RR']:
        try:
            n = doc.Tree.FindNode(p)
            vals[p.split('\\')[-1]] = n.Value if n is not None else None
        except Exception as e:
            vals[p.split('\\')[-1]] = 'ERR'
    print('   [%s] %s  Ready=%r' % (tag, vals, doc.Engine.Ready), flush=True)


doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(REF)
time.sleep(4)
read_out(doc, '刚打开')

print('   改 BASIS_RR: 24 -> 20', flush=True)
try:
    n = doc.Tree.FindNode(r'\Data\Blocks\B1\Input\BASIS_RR')
    n.Value = 20.0
    print('   写入后 =', doc.Tree.FindNode(r'\Data\Blocks\B1\Input\BASIS_RR').Value, flush=True)
except Exception as e:
    print('   写入失败', str(e)[:100], flush=True)

print('   运行中...', flush=True)
t0 = time.time()
try:
    doc.Engine.Run2()
except Exception as e:
    print('   Run2 异常', str(e)[:100], flush=True)
while time.time() - t0 < 60:
    time.sleep(3)
    try:
        if doc.Engine.IsRunning is False:
            break
    except Exception:
        break
print('   用时 %.0fs' % (time.time() - t0), flush=True)
read_out(doc, '运行后')

# 输出字段名全列
n = doc.Tree.FindNode(r'\Data\Blocks\B1\Output')
names = [n.Elements.Item(i).Name for i in range(n.Elements.Count)]
print()
print('   B1.Output 全部字段:', [x for x in names if 'DUTY' in x or 'RR' in x or 'TEMP' in x], flush=True)
try:
    doc.Close()
except Exception:
    pass
