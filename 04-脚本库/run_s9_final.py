# -*- coding: utf-8 -*-
"""定稿：列出全部设备与连接、关键流股表，另存 apwz"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32

doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
doc.InitFromArchive2(r'D:\<化工工作区>\NA-Chemical-10000t_s8_s.bkp')
time.sleep(3)

L = []
nb = doc.Tree.FindNode(r'\Data\Blocks')
names = []
for i in range(nb.Elements.Count):
    try:
        names.append(nb.Elements.Item(i).Name)
    except Exception:
        pass
L.append('=== 设备清单（%d 台）===' % len(names))
for b in sorted(names):
    t = None
    try:
        t = doc.Tree.FindNode(r'\Data\Blocks\%s\Input\TYPE' % b).Value
    except Exception:
        pass
    L.append('  %-8s %s' % (b, t))

L.append('')
ns = doc.Tree.FindNode(r'\Data\Streams')
sn = []
for i in range(ns.Elements.Count):
    try:
        sn.append(ns.Elements.Item(i).Name)
    except Exception:
        pass
L.append('=== 流股（%d 条）===' % len(sn))
L.append('  ' + ' '.join(sorted(sn)))

L.append('')
L.append('=== 流股温度/压力/质量流量 ===')
L.append('  %-8s %-10s %-10s %-12s' % ('流股', 'T(℃)', 'P(bar)', 'W(kg/h)'))
for s in sorted(sn):
    def gg(k):
        try:
            return doc.Tree.FindNode(r'\Data\Streams\%s\Output\%s\MIXED' % (s, k)).Value
        except Exception:
            return None
    L.append('  %-8s %-10s %-10s %-12s' % (s, gg('TEMP'), gg('PRES'), gg('MASSFLMX')))

open(r'D:\<化工工作区>\_probe\final_dump.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('设备 %d 台，流股 %d 条' % (len(names), len(sn)))
for target in [r'D:\<化工工作区>\NA-Chemical-10000t_定稿.apwz', r'D:\<化工工作区>\NA-Chemical-10000t_定稿.bkp']:
    try:
        doc.SaveAs(target)
        print('已另存', target)
    except Exception as ex:
        print('另存失败', target, ex)
try:
    doc.Close()
except Exception:
    pass
print('DONE')
