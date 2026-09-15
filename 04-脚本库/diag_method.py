# -*- coding: utf-8 -*-
"""诊断：换物性方法试探 AE_UNDERSPEC 的真正原因 + 读引擎消息"""
import time
import win32com.client as win32

F = r'D:\<化工工作区>\NA-Chemical-10000t_模拟.bkp'
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(F)
time.sleep(4)

print('=== Engine 可用属性 ===')
print([m for m in dir(doc.Engine) if not m.startswith('_')][:40], flush=True)
print()

print('=== Results Summary 子树 ===')
rs = doc.Tree.FindNode(r'\Data\Results Summary')
if rs is not None:
    for i in range(rs.Elements.Count):
        ch = rs.Elements.Item(i)
        print('   %-24s 子=%s' % (ch.Name, ch.Elements.Count if hasattr(ch, 'Elements') else '-'), flush=True)
print()

print('=== 换物性方法试探 ===')
for method in ['PENG-ROB', 'IDEAL', 'NRTL']:
    p = r'\Data\Properties\Specifications\Input\GOPSETNAME'
    try:
        doc.Tree.FindNode(p).Value = method
        got = doc.Tree.FindNode(p).Value
    except Exception as e:
        print('  设 %-10s 失败 %s' % (method, str(e)[:60]), flush=True)
        continue
    try:
        doc.Tree.FindNode(r'\Data\Streams\S-101\Input\TEMP').Value = 25.0
        print('  %-10s (读到 %r) -> TEMP 写入成功!' % (method, got), flush=True)
        break
    except Exception as e:
        # 只取错误码部分
        s = str(e)
        code = s.split("'")[3] if s.count("'") > 3 else s[:60]
        print('  %-10s (读到 %r) -> TEMP 仍失败: %s' % (method, got, code), flush=True)

print()
print('=== 换个流股试（S-107 纯水，最单纯）===')
for s in ['S-107', 'S-111', 'S-124']:
    try:
        doc.Tree.FindNode(r'\Data\Streams\%s\Input\TEMP' % s).Value = 25.0
        print('  %-8s TEMP 成功' % s, flush=True)
    except Exception as e:
        print('  %-8s TEMP 失败 %s' % (s, str(e)[:70]), flush=True)
try:
    doc.Close()
except Exception:
    pass
