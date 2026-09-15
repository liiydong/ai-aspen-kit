# -*- coding: utf-8 -*-
"""打开真实 bkp 副本，摸清树路径（组分/模块/流股）"""
import time, sys
import win32com.client as win32

doc = win32.Dispatch('Apwn.Document')
p = r'D:\<化工工作区>\_probe\probe_copy.apwz'
print('打开:', p)
sys.stdout.flush()
try:
    doc.InitFromArchive2(p)
except Exception as e:
    print('InitFromArchive2 失败:', repr(e)[:200])
time.sleep(3)

def dump(path, label, limit=30):
    print(f'\n=== {label} :: {path}')
    try:
        n = doc.Tree.FindNode(path)
        if n is None:
            print('  节点为 None'); return
        print('  Name =', n.Name, '| 子元素 =', n.Elements.Count)
        for i in range(min(limit, n.Elements.Count)):
            try:
                print('    ', i, n.Elements.Item(i).Name)
            except Exception as e:
                print('    ', i, 'err', str(e)[:60])
    except Exception as e:
        print('  失败:', str(e)[:160])

dump(r'\Data\Components\Specifications\Input', 'Components.Input', 20)
dump(r'\Data\Blocks', 'Blocks', 30)
dump(r'\Data\Streams', 'Streams', 30)
dump(r'\Data\Properties\Specifications\Input', 'Properties.Input', 15)

print('\n=== 尝试读一个流股的值 ===')
try:
    st = doc.Tree.FindNode(r'\Data\Streams')
    if st.Elements.Count:
        name = st.Elements.Item(0).Name
        print('第一个流股:', name)
        for sub in [r'\Input\TEMP', r'\Input\PRES', r'\Input\MASSFLMX']:
            try:
                v = doc.Tree.FindNode(r'\Data\Streams\\' + name + sub).Value
                print('  ', sub, '=', v)
            except Exception as e:
                print('  ', sub, 'err', str(e)[:70])
except Exception as e:
    print('读流股失败:', str(e)[:150])
print('done')
