# -*- coding: utf-8 -*-
"""在完整树下重测：能否创建组分、流股、模块"""
import time
import win32com.client as win32

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitNew2()
time.sleep(2)

for p in [r'\Data', r'\Data\Components\Specifications\Input',
          r'\Data\Streams', r'\Data\Blocks']:
    n = doc.Tree.FindNode(p)
    print('%-42s %s' % (p, 'None' if n is None else '存在(子=%s)' % n.Elements.Count))
print()

TRIALS = [
    (r'\Data\Components\Specifications\Input', 'Add("98-92-0")', lambda n: n.Elements.Add('98-92-0')),
    (r'\Data\Components\Specifications\Input', 'Add("NICOTINAMIDE")', lambda n: n.Elements.Add('NICOTINAMIDE')),
    (r'\Data\Components\Specifications\Input', 'Add("NAM","CONV")', lambda n: n.Elements.Add('NAM', 'CONV')),
    (r'\Data\Streams', 'Add("S-101")', lambda n: n.Elements.Add('S-101')),
    (r'\Data\Blocks', 'Add("M-101","Mixer")', lambda n: n.Elements.Add('M-101', 'Mixer')),
    (r'\Data\Blocks', 'Add("M-101","Mixer","Mixer")', lambda n: n.Elements.Add('M-101', 'Mixer', 'Mixer')),
]

for path, desc, fn in TRIALS:
    n = doc.Tree.FindNode(path)
    if n is None:
        print('%-34s 节点不存在' % desc)
        continue
    try:
        r = fn(n)
        print('%-34s 成功 返回=%r' % (desc, r))
    except Exception as e:
        print('%-34s 失败 %s' % (desc, str(e)[:90]))

print()
# 看看各节点的 Elements 支持什么
for path in [r'\Data\Components\Specifications\Input', r'\Data\Streams', r'\Data\Blocks']:
    n = doc.Tree.FindNode(path)
    if n is None:
        continue
    el = n.Elements
    print('%-42s Elements 属性: %s' % (path, [a for a in dir(el) if not a.startswith('_')]))
try:
    doc.Close()
except Exception:
    pass
