# -*- coding: utf-8 -*-
"""拿报错清单：ExportReport 各种签名 + 运行后看模块输出"""
import time, os, glob
import win32com.client as win32

F = r'D:\<化工工作区>\NA-Chemical-10000t_模拟.bkp'
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(F)
time.sleep(4)
print('Ready =', doc.Engine.Ready, flush=True)

print()
print('=== ExportReport 签名试探 ===', flush=True)
OUT = r'D:\<化工工作区>\_probe\cp'
for args in [(OUT + '.txt',), (1, OUT + '1.txt'), (0, OUT + '0.txt'),
             (OUT + '2.txt', 1), (2, OUT + '3.txt')]:
    try:
        doc.Engine.ExportReport(*args)
        print('  参数 %d 个 -> 成功 %s' % (len(args), args), flush=True)
    except Exception as e:
        print('  参数 %d 个 -> 失败 %s' % (len(args), str(e)[:70]), flush=True)
for f in glob.glob(OUT + '*.txt'):
    print('  生成文件:', f, os.path.getsize(f), flush=True)
    print(open(f, encoding='utf-8', errors='ignore').read()[:1500], flush=True)

print()
print('=== 运行并列模块输出 ===', flush=True)
try:
    doc.Engine.Run2()
except Exception as e:
    print('  Run2 异常', str(e)[:100], flush=True)
time.sleep(8)
for b in ['M-101', 'E-101', 'R-101', 'T-201']:
    n = doc.Tree.FindNode(r'\Data\Blocks\%s\Output' % b)
    if n is None:
        print('  %s: Output 节点无' % b, flush=True)
        continue
    names = [n.Elements.Item(i).Name for i in range(n.Elements.Count)]
    print('  %s Output 字段数 %d: %s' % (b, len(names), names[:14]), flush=True)
    for f in ['TEMP', 'PRES', 'DUTY', 'B-PRES']:
        try:
            x = n.Elements.Item(f)
            print('      %s = %r' % (f, x.Value), flush=True)
        except Exception:
            pass
try:
    doc.Close()
except Exception:
    pass
