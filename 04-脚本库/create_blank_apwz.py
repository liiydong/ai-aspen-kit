# -*- coding: utf-8 -*-
"""尝试用 COM 新建并保存一个空白 Aspen 文件"""
import time, os, sys
import win32com.client as win32

target = r'D:\<化工工作区>\NA-Chemical-10000t.apwz'
if os.path.exists(target):
    print('文件已存在，跳过创建:', target); sys.exit(0)

doc = win32.Dispatch('Apwn.Document')
doc.InitNew()
time.sleep(3)
print('InitNew 完成')

ok = False
for m in ('SaveAs', 'Save', 'SaveAs2', 'SaveAsFile'):
    fn = getattr(doc, m, None)
    if fn is None:
        print(f'  {m}: 该方法不存在'); continue
    try:
        fn(target)
        print(f'  {m}: 调用成功')
        ok = True
        break
    except Exception as e:
        print(f'  {m}: 失败 {str(e)[:110]}')

time.sleep(2)
print('文件是否生成:', os.path.exists(target))
if ok and os.path.exists(target):
    print('大小:', os.path.getsize(target), 'bytes')
