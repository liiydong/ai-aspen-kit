# -*- coding: utf-8 -*-
"""加载后尝试 Reinit / Run2 / Engine.Reinit，看流程图能否被激活"""
import time
import win32com.client as win32

BKP = r'D:\<化工工作区>\_probe\base.bkp'
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(BKP)
time.sleep(3)


def show(tag):
    try:
        d = doc.Tree.FindNode(r'\Data')
        s = doc.Tree.FindNode(r'\Data\Streams')
        b = doc.Tree.FindNode(r'\Data\Blocks')
        print('   %-22s \\Data=%s Streams=%s Blocks=%s' % (
            tag, d.Elements.Count if d else 0,
            'None' if s is None else s.Elements.Count,
            'None' if b is None else b.Elements.Count), flush=True)
    except Exception as e:
        print('   %-22s 读取失败 %s' % (tag, str(e)[:60]), flush=True)


show('初始')

for name, fn in [
    ('doc.Reinit()', lambda: doc.Reinit()),
    ('Engine.Reinit()', lambda: doc.Engine.Reinit()),
    ('Engine.Check()', lambda: doc.Engine.Check()),
    ('Engine.Run2()', lambda: doc.Engine.Run2()),
]:
    try:
        fn()
        print('  >> %s 调用成功' % name, flush=True)
    except Exception as e:
        print('  >> %s 失败 %s' % (name, str(e)[:90]), flush=True)
    time.sleep(2)
    show(name)

try:
    doc.Close()
except Exception:
    pass
