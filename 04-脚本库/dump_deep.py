# -*- coding: utf-8 -*-
"""深度遍历 Aspen 树，输出到文件"""
import time, sys
import win32com.client as win32

P = sys.argv[1] if len(sys.argv) > 1 else r'D:\<化工工作区>\_probe\base.bkp'
OUT = r'D:\<化工工作区>\_probe\tree_dump.txt'

doc = win32.DispatchEx('Apwn.Document')
doc.InitFromArchive2(P)
time.sleep(3)

lines = []
seen = set()


def walk(node, path, depth, maxdepth):
    if depth > maxdepth:
        return
    try:
        cnt = node.Elements.Count
    except Exception:
        return
    try:
        val = node.Value
    except Exception:
        val = None
    lines.append('%s%s  [%d]  %r' % ('  ' * depth, path, cnt, val))
    for i in range(cnt):
        try:
            ch = node.Elements.Item(i)
        except Exception:
            continue
        nm = None
        try:
            nm = ch.Name
        except Exception:
            pass
        if nm is None:
            continue
        walk(ch, path + '\\' + nm, depth + 1, maxdepth)


for top in ['Data']:
    n = doc.Tree.FindNode('\\' + top)
    if n is None:
        lines.append('!! %s 不存在' % top)
    else:
        walk(n, '\\' + top, 0, 3)

open(OUT, 'w', encoding='utf-8').write('\n'.join(lines))
print('已写', OUT, '行数', len(lines))
print()
for l in lines[:120]:
    print(l)
try:
    doc.Close()
except Exception:
    pass
