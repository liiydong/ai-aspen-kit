# -*- coding: utf-8 -*-
"""用 DispatchEx 强制新建实例，打开用户搭建的文件"""
import time, sys
import win32com.client as win32

TARGET = r'D:\<化工工作区>\NA-Chemical-10000t.apwz'
print('新建独立 Aspen 实例 ...')
sys.stdout.flush()
doc = win32.DispatchEx('Apwn.Document')
print('实例已创建')
sys.stdout.flush()
try:
    doc.InitFromArchive2(TARGET)
    print('>>> 打开成功:', TARGET)
except Exception as e:
    print('>>> 打开失败:', str(e)[:200])
    sys.exit(1)

time.sleep(3)

def lst(path, label, limit=50):
    try:
        n = doc.Tree.FindNode(path)
        c = n.Elements.Count
        names = [n.Elements.Item(i).Name for i in range(min(limit, c))]
        print(f'{label} 共 {c} 个:')
        print('   ', ', '.join(names))
        return names
    except Exception as e:
        print(f'{label} 读取失败:', str(e)[:150])
        return []

print()
blocks = lst(r'\Data\Blocks', '模块 Blocks')
streams = lst(r'\Data\Streams', '流股 Streams')

print('\n--- 模块类型 ---')
for b in blocks:
    for suffix in (r'\Input\MODEL', r'\Input\TYPE'):
        try:
            v = doc.Tree.FindNode(r'\Data\Blocks\\' + b + suffix).Value
            if v:
                print(f'  {b} = {v}')
                break
        except Exception:
            pass
