# -*- coding: utf-8 -*-
"""跑 full3 + 用参考模型确定流股输出字段名"""
import time
import win32com.client as win32

# 1) 先从参考模型拿流股 Output 的正确字段名
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(r'D:\<化工工作区>\_probe\ref_multi.bkp')
time.sleep(4)
n = doc.Tree.FindNode(r'\Data\Streams\S5\Output')
names = [n.Elements.Item(i).Name for i in range(n.Elements.Count)]
print('参考模型 S5.Output 字段(%d):' % len(names), flush=True)
print('  ', [x for x in names if any(k in x.upper() for k in
      ['TEMP', 'PRES', 'FLOW', 'MASS', 'MOLE', 'VAPOR', 'FRAC'])][:20], flush=True)
for f in names:
    if f in ('TEMP', 'PRES', 'VFRAC', 'MASSFLMX'):
        try:
            print('   %s = %r' % (f, n.Elements.Item(f).Value), flush=True)
        except Exception as e:
            print('   %s 读取失败 %s' % (f, str(e)[:50]), flush=True)
print('  原始子节点数:', n.Elements.Count, flush=True)
print('  TEMP 直接读:', end=' ', flush=True)
try:
    print(repr(n.Elements.Item('TEMP').Value), flush=True)
except Exception as e:
    print('失败', str(e)[:60], flush=True)
try:
    doc.Close()
except Exception:
    pass

# 2) 跑我们的模型
print()
print('=' * 20, '跑我们的模型', flush=True)
doc2 = win32.DispatchEx('Apwn.Document')
try:
    doc2.SuppressDialogs = True
except Exception:
    pass
doc2.InitFromArchive(r'D:\<化工工作区>\_probe\full3.bkp')
time.sleep(4)
print('  Ready =', doc2.Engine.Ready, flush=True)
t0 = time.time()
try:
    doc2.Engine.Run2()
    print('  Run2 完成 %.1fs' % (time.time() - t0), flush=True)
except Exception as e:
    print('  Run2 异常', str(e)[:120], flush=True)
for i in range(15):
    time.sleep(3)
    try:
        if doc2.Engine.IsRunning is False:
            break
    except Exception:
        break
print()
for b in ['M-101', 'E-101', 'R-101', 'T-201', 'R-501', 'T-404']:
    nn = doc2.Tree.FindNode(r'\Data\Blocks\%s\Output' % b)
    if nn is None:
        print('  %s: 无 Output' % b, flush=True)
        continue
    vals = []
    for f in ['TEMP', 'PRES', 'DUTY', 'B-PRES']:
        try:
            vals.append('%s=%s' % (f, nn.Elements.Item(f).Value))
        except Exception:
            pass
    print('  %-7s %s' % (b, ' '.join(vals)), flush=True)
try:
    doc2.SaveAs(r'D:\<化工工作区>\_probe\full3_saved.bkp')
    print('  已保存 full3_saved.bkp', flush=True)
except Exception:
    pass
try:
    doc2.Close()
except Exception:
    pass
