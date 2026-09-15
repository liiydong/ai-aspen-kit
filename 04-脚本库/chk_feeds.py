# -*- coding: utf-8 -*-
import sys, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_sim5.bkp'
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive2(BASE)
time.sleep(3)

print('=== \\Data 子节点 ===')
n = doc.Tree.FindNode(r'\Data')
for i in range(n.Elements.Count):
    try:
        c = n.Elements.Item(i)
        print('   ', c.Name)
    except Exception:
        pass

print()
print('=== 流股输入 ===')
for s in ['S-101', 'S-102', 'S-103', 'S-106', 'S-107', 'S-108', 'S-109', 'S-110',
          'S-111', 'S-112', 'S-114', 'S-123', 'S-124', 'S-125']:
    p = r'\Data\Streams\%s\Input' % s
    nd = doc.Tree.FindNode(p)
    if nd is None:
        print('  %-6s 无输入节点' % s); continue
    got = []
    for f in ['TEMP', 'PRES', 'MOLE-FLOW', 'MASS-FLOW', 'VOLUME-FLOW', 'TOTAL']:
        try:
            e = nd.Elements.Item(f)
            if e is not None and e.Value not in (None, ''):
                got.append('%s=%s' % (f, e.Value))
        except Exception:
            pass
    # 组分流量
    try:
        mx = nd.Elements.Item('MIXED')
        if mx is not None:
            cnt = mx.Elements.Count
            got.append('MIXED子项=%d' % cnt)
            showed = 0
            for i in range(cnt):
                try:
                    cc = mx.Elements.Item(i)
                    if cc.Value not in (None, '', 0, 0.0) and showed < 6:
                        got.append('%s=%s' % (cc.Name, cc.Value)); showed += 1
                except Exception:
                    pass
    except Exception:
        pass
    print('  %-6s %s' % (s, ', '.join(got) if got else '(全空)'))

print()
print('=== 关键流股输出（上一轮计算结果） ===')
for s in ['S-106', 'S-108', 'S-110', 'S-112']:
    nd = doc.Tree.FindNode(r'\Data\Streams\%s\Output' % s)
    if nd is None:
        print('  %-6s 无' % s); continue
    got = []
    for f in ['TEMP', 'PRES', 'MOLE-FLOW', 'MASS-FLOW']:
        try:
            e = nd.Elements.Item(f)
            if e is not None and e.Value not in (None, ''):
                got.append('%s=%s' % (f, e.Value))
        except Exception:
            pass
    print('  %-6s %s' % (s, ', '.join(got) if got else '(空)'))

try:
    doc.Close()
except Exception:
    pass
print('done')
