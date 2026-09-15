# -*- coding: utf-8 -*-
"""探测 Aspen COM 能否新增组分/模块（决定自动化边界）"""
import time, sys
import win32com.client as win32

def show(title, fn):
    try:
        print(f'[{title}]', fn())
    except Exception as e:
        print(f'[{title}] 失败:', repr(e)[:200])

doc = win32.Dispatch('Apwn.Document')
print('已连接。尝试 InitNew ...')
sys.stdout.flush()
try:
    doc.InitNew()
    print('InitNew 成功')
except Exception as e:
    print('InitNew 失败:', repr(e)[:200])

time.sleep(3)

# 1) 组分节点
try:
    comp = doc.Tree.FindNode(r'\Data\Components\Specifications\Input')
    print('组分组件元素数:', comp.Elements.Count)
    for i in range(min(3, comp.Elements.Count)):
        el = comp.Elements.Item(i)
        print('   子节点', i, '=', el.Name)
    try:
        comp.Elements.Add()
        print('>>> 组分 Elements.Add() 可用，新数量:', comp.Elements.Count)
    except Exception as e:
        print('>>> 组分 Elements.Add() 不可用:', repr(e)[:150])
except Exception as e:
    print('组分节点访问失败:', repr(e)[:200])

# 2) 模块节点
try:
    blk = doc.Tree.FindNode(r'\Data\Blocks')
    print('Blocks 元素数:', blk.Elements.Count)
    try:
        blk.Elements.Add('TESTB1')
        print('>>> Blocks.Elements.Add(名) 可用，新数量:', blk.Elements.Count)
    except Exception as e:
        print('>>> Blocks.Elements.Add 不可用:', repr(e)[:150])
except Exception as e:
    print('Blocks 节点访问失败:', repr(e)[:200])

# 3) 流股节点
try:
    st = doc.Tree.FindNode(r'\Data\Streams')
    print('Streams 元素数:', st.Elements.Count)
except Exception as e:
    print('Streams 节点访问失败:', repr(e)[:200])

print('探测结束')
