# -*- coding: utf-8 -*-
"""试 Tree.NewChild / Elements.Add 加组分，并验证二元参数检索机制"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>'
doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
doc.InitFromArchive2(os.path.join(BASE, 'NA_work.bkp'))
t = time.time()
while time.time() - t < 25:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)

out = []

# --- A) 看组分表结构 ---
p = r'\Data\Components\Specifications\Input'
nd = doc.Tree.FindNode(p)
out.append('Specifications\\Input: %s' % ('None' if nd is None else 'OK'))
if nd is not None:
    try:
        out.append('  Elements.Count = %d' % nd.Elements.Count)
        for i in range(min(nd.Elements.Count, 20)):
            e = nd.Elements.Item(i)
            try:
                nm = e.Name
            except Exception:
                nm = '?'
            try:
                cc = e.Elements.Count
            except Exception:
                cc = '-'
            out.append('    %-14s rows=%s' % (nm, cc))
    except Exception as ex:
        out.append('  ' + str(ex)[:90])

# --- B) 试往 ANAME 列加行 ---
out.append('')
out.append('--- 试 ANAME.Elements.Add("METH") ---')
nd = doc.Tree.FindNode(r'\Data\Components\Specifications\Input\ANAME')
before = None
if nd is not None:
    try:
        before = nd.Elements.Count
    except Exception:
        pass
    out.append('  before rows = %s' % before)
    try:
        nd.Elements.Add('METH')
        out.append('  Add 成功')
    except Exception as ex:
        out.append('  Add 失败: %s' % str(ex)[:120])
    try:
        out.append('  after rows = %s' % nd.Elements.Count)
    except Exception as ex:
        out.append('  ' + str(ex)[:80])

# --- C) 试 Tree.NewChild ---
out.append('')
out.append('--- 试 Tree.NewChild ---')
for args in [(r'\Data\Components\Specifications\Input\ANAME', 'METH2'),
             (r'\Data\Components', 'METH3')]:
    try:
        r = doc.Tree.NewChild(*args)
        out.append('  NewChild%s 成功 -> %r' % (args, r))
    except Exception as ex:
        out.append('  NewChild%s 失败: %s' % (args, str(ex)[:110]))

# --- D) 观察组分总数变化 ---
out.append('')
out.append('--- Components\\Specifications\\Input\\OUTNAME 行数 ---')
nd = doc.Tree.FindNode(r'\Data\Components\Specifications\Input\OUTNAME')
if nd is not None:
    try:
        out.append('  rows = %d' % nd.Elements.Count)
        names = []
        for i in range(nd.Elements.Count):
            try:
                names.append(nd.Elements.Item(i).Name)
            except Exception:
                pass
        out.append('  ' + ', '.join(names))
    except Exception as ex:
        out.append('  ' + str(ex)[:90])

print('\n'.join(out))
try:
    doc.Close()
except Exception:
    pass
print('DONE')
