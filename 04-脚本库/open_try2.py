# -*- coding: utf-8 -*-
"""用副本 + 多种打开方法测试"""
import shutil, time, sys, os
import win32com.client as win32

src = r'D:\<化工工作区>\NA-Chemical-10000t.apwz'
dst = r'D:\<化工工作区>\_probe\na_copy2.apwz'
os.makedirs(os.path.dirname(dst), exist_ok=True)
shutil.copy2(src, dst)
print('副本已建:', dst, os.path.getsize(dst), 'bytes')

doc = win32.DispatchEx('Apwn.Document')
print('实例已创建')
sys.stdout.flush()
for m in ('InitFromArchive2', 'InitFromArchive', 'InitFromFile'):
    fn = getattr(doc, m, None)
    if fn is None:
        print(f'  {m}: 方法不存在'); continue
    try:
        fn(dst)
        print(f'  >>> {m} 成功打开副本')
        time.sleep(3)
        n = doc.Tree.FindNode(r'\Data\Blocks')
        print('      模块数 =', n.Elements.Count)
        break
    except Exception as e:
        print(f'  {m} 失败: {str(e)[:130]}')

# 若都不行，尝试 InitNew 直接新建
print('\n--- 尝试 InitNew（验证 COM 本身是否可用） ---')
try:
    doc2 = win32.DispatchEx('Apwn.Document')
    doc2.InitNew()
    time.sleep(2)
    print('InitNew 成功；Blocks 数 =', doc2.Tree.FindNode(r'\Data\Blocks').Elements.Count)
except Exception as e:
    print('InitNew 失败:', str(e)[:150])
