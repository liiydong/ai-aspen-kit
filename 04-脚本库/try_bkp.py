# -*- coding: utf-8 -*-
"""从 apwz 中解出 bkp，再尝试打开"""
import zipfile, os, time, sys
import win32com.client as win32

src = r'D:\<化工工作区>\NA-Chemical-10000t.apwz'
out = r'D:\<化工工作区>\_probe\na_from_apwz.bkp'
with zipfile.ZipFile(src) as z:
    name = z.namelist()[0]
    data = z.read(name)
with open(out, 'wb') as f:
    f.write(data)
print('已解出 bkp:', out, '|', len(data), 'bytes')

doc = win32.DispatchEx('Apwn.Document')
print('实例已创建')
sys.stdout.flush()
ok = False
for m in ('InitFromArchive2', 'InitFromArchive', 'InitFromFile'):
    fn = getattr(doc, m, None)
    if fn is None:
        continue
    try:
        fn(out)
        print(f'>>> {m} 成功打开 bkp')
        ok = True
        break
    except Exception as e:
        print(f'  {m} 失败: {str(e)[:130]}')

if ok:
    time.sleep(3)
    try:
        b = doc.Tree.FindNode(r'\Data\Blocks')
        s = doc.Tree.FindNode(r'\Data\Streams')
        print('模块数:', b.Elements.Count)
        print('流股数:', s.Elements.Count)
        print('模块:', ', '.join(b.Elements.Item(i).Name for i in range(b.Elements.Count)))
        print('流股:', ', '.join(s.Elements.Item(i).Name for i in range(s.Elements.Count)))
    except Exception as e:
        print('读拓扑失败:', str(e)[:150])
