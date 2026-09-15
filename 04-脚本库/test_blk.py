# -*- coding: utf-8 -*-
"""攻模块创建签名 + 组分表格 InsertRow"""
import time
import win32com.client as win32

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitNew2()
time.sleep(2)

# ---------- 1) 模块创建签名 ----------
blk = doc.Tree.FindNode(r'\Data\Blocks')
print('== 模块创建签名尝试 ==')
CANDS = [
    ('Add("M-101")', ('M-101',)),
    ('Add("M-101","MIXER")', ('M-101', 'MIXER')),
    ('Add("M-101","Mixer")', ('M-101', 'Mixer')),
    ('Add("M-101","MIXER","Mixer")', ('M-101', 'MIXER', 'Mixer')),
    ('Add("M-101","Mixer","")', ('M-101', 'Mixer', '')),
    ('Add("M-101","","Mixer")', ('M-101', '', 'Mixer')),
    ('Add("M-101","Mixer","AspenTech")', ('M-101', 'Mixer', 'AspenTech')),
    ('Add("M-101","Mixer","Built-In")', ('M-101', 'Mixer', 'Built-In')),
]
for desc, args in CANDS:
    try:
        blk.Elements.Add(*args)
        print('  %-38s 成功!  现有模块=%s' % (desc, blk.Elements.Count))
    except Exception as e:
        print('  %-38s 失败 %s' % (desc, str(e)[:80]))

print()
# ---------- 2) 组分表格 ----------
c = doc.Tree.FindNode(r'\Data\Components\Specifications\Input')
el = c.Elements
print('== 组分表格 ==')
for a in ['Count', 'RowCount', 'Dimension', 'DimensionName', 'Label', 'LabelLocation', 'ItemName']:
    try:
        print('  %-14s = %r' % (a, getattr(el, a)))
    except Exception as e:
        print('  %-14s 读取失败 %s' % (a, str(e)[:60]))
try:
    r = el.Item(0)
    print('  Item(0).Name =', r.Name)
    print('  Item(0) 子节点 =', [r.Elements.Item(i).Name for i in range(r.Elements.Count)])
except Exception as e:
    print('  Item(0) 失败:', str(e)[:80])
try:
    el.InsertRow()
    print('  InsertRow() 成功, RowCount =', el.RowCount)
except Exception as e:
    print('  InsertRow() 失败:', str(e)[:90])
try:
    el.InsertRow(0)
    print('  InsertRow(0) 成功, RowCount =', el.RowCount)
except Exception as e:
    print('  InsertRow(0) 失败:', str(e)[:90])
try:
    print('  Values =', el.Values)
except Exception as e:
    print('  Values 失败:', str(e)[:80])
try:
    doc.Close()
except Exception:
    pass
