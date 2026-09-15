# -*- coding: utf-8 -*-
"""打开用户搭好的 Aspen 文件并列出拓扑"""
import time, sys
import win32com.client as win32

TARGET = r'D:\<化工工作区>\NA-Chemical-10000t.apwz'
doc = win32.Dispatch('Apwn.Document')
print('尝试打开:', TARGET)
sys.stdout.flush()
try:
    doc.InitFromArchive2(TARGET)
    print('打开成功')
except Exception as e:
    print('InitFromArchive2 失败:', str(e)[:200])
    print('退出码 1')
    sys.exit(1)

time.sleep(3)

def count(path, label):
    try:
        n = doc.Tree.FindNode(path)
        c = n.Elements.Count
        names = [n.Elements.Item(i).Name for i in range(c)]
        print(f'{label}: {c} 个')
        print('   ', ', '.join(names))
        return names
    except Exception as e:
        print(f'{label} 读取失败:', str(e)[:160])
        return []

print()
blocks = count(r'\Data\Blocks', '模块 Blocks')
count(r'\Data\Streams', '流股 Streams')

# 组分：从 Setup 或 Components 读
for p in [r'\Data\Components\Specifications\Input\Components',
          r'\Data\Components\Specifications\Input']:
    try:
        n = doc.Tree.FindNode(p)
        print(f'组分节点 {p}: 子元素 {n.Elements.Count}')
    except Exception as e:
        print(f'组分节点 {p} 失败: {str(e)[:120]}')

print('\n--- 检查每个模块的模型类型 ---')
for b in blocks[:20]:
    for mp in [r'\Data\Blocks' + '\\' + b + r'\Input\MODEL',
               r'\Data\Blocks' + '\\' + b + r'\Input\TYPE']:
        try:
            v = doc.Tree.FindNode(mp).Value
            if v:
                print(f'  {b}: {v}')
                break
        except Exception:
            pass
