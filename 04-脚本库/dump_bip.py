# -*- coding: utf-8 -*-
"""dump BDBANK 版文件里 Aspen 实际检索到的 NRTL 二元参数"""
import sys, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

P = r'D:\<化工工作区>\_probe\bipA2.bkp'
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive2(P)
time.sleep(4)


def walk(node, path, depth, out):
    if depth > 4:
        return
    try:
        cnt = node.Elements.Count
    except Exception:
        return
    for k in range(cnt):
        try:
            e = node.Elements.Item(k)
        except Exception:
            continue
        nm = None
        try:
            nm = e.Name
        except Exception:
            pass
        if nm is None:
            continue
        np_ = path + '\\' + nm
        out.append((depth, np_))
        walk(e, np_, depth + 1, out)


base = doc.Tree.FindNode(r'\Data\Properties\Parameters\Binary Interaction\NRTL-1')
print('NRTL-1 节点:', base is not None)
if base is not None:
    try:
        print('  子节点:', [base.Elements.Item(k).Name for k in range(base.Elements.Count)])
    except Exception as ex:
        print('  ', ex)
    for sub in ['Input', 'Output']:
        n2 = doc.Tree.FindNode(r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\%s' % sub)
        if n2 is None:
            print('  %s: None' % sub)
            continue
        try:
            ch = [n2.Elements.Item(k).Name for k in range(n2.Elements.Count)]
            print('  %s 子节点(%d): %s' % (sub, len(ch), ch[:20]))
        except Exception as ex:
            print('  %s:' % sub, ex)
        # 逐个子节点找表
        for k in range(min(12, n2.Elements.Count)):
            try:
                e = n2.Elements.Item(k)
                en = e.Name
            except Exception:
                continue
            try:
                v = e.Value
            except Exception:
                v = None
            print('     %-28s value=%s elements=%s' % (
                en, v, getattr(e.Elements, 'Count', '?')))

# 完整 dump
print()
print('=== 全路径 dump ===')
out = []
walk(doc.Tree.FindNode(r'\Data\Properties\Parameters'), r'\Data\Properties\Parameters', 0, out)
for d, p in out:
    if 'NRTL' in p.upper() or 'BPVAL' in p.upper() or d <= 2:
        print('  ' * d, p)
print()
print('总节点数:', len(out))
try:
    doc.Close()
except Exception:
    pass
print('DONE')
