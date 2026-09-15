# -*- coding: utf-8 -*-
"""对照：参考模型能否写流股 TEMP"""
import time
import win32com.client as win32

REF = r'D:\<化工工作区>\_probe\ref_multi.bkp'
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(REF)
time.sleep(4)

st = doc.Tree.FindNode(r'\Data\Streams')
names = [st.Elements.Item(i).Name for i in range(st.Elements.Count)]
print('参考模型流股:', names, flush=True)

n = doc.Tree.FindNode(r'\Data\Streams\S5\Input')
print()
print('S5 Input 中含 TEMP/PRES/BASIS/SPEC 的字段:')
for i in range(n.Elements.Count):
    ch = n.Elements.Item(i)
    if any(k in ch.Name.upper() for k in ['TEMP', 'PRES', 'BASIS', 'SPEC', 'FLASH']):
        try:
            v = repr(ch.Value)[:40]
        except Exception:
            v = ''
        print('   %-18s %s' % (ch.Name, v), flush=True)

print()
print('=== 试写 S5 的 TEMP（原值 95.22）===')
for s in ['S5', 'S4', 'S9']:
    for f, v in [('TEMP', 95.2213461), ('PRES', 0.221034483)]:
        try:
            doc.Tree.FindNode(r'\Data\Streams\%s\Input\%s' % (s, f)).Value = v
            print('   %s.%s 成功' % (s, f), flush=True)
        except Exception as e:
            print('   %s.%s 失败 %s' % (s, f, str(e)[:70]), flush=True)

print()
print('=== 对照：改成别的值再试 ===')
try:
    doc.Tree.FindNode(r'\Data\Streams\S5\Input\TEMP').Value = 80.0
    print('   S5.TEMP = 80 成功 -> %r' % doc.Tree.FindNode(r'\Data\Streams\S5\Input\TEMP').Value, flush=True)
except Exception as e:
    print('   S5.TEMP = 80 失败', str(e)[:80], flush=True)
try:
    doc.Close()
except Exception:
    pass
