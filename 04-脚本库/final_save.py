# -*- coding: utf-8 -*-
"""定稿：另存 apwz，并确认设备/流股数"""
import sys, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32

doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
doc.InitFromArchive2(r'D:\<化工工作区>\NA-Chemical-10000t_s13_s.bkp')
time.sleep(3)
nb = doc.Tree.FindNode(r'\Data\Blocks')
ns = doc.Tree.FindNode(r'\Data\Streams')
print('设备 %d 台；流股 %d 条' % (nb.Elements.Count, ns.Elements.Count))
print('设备:', ' '.join(sorted(nb.Elements.Item(i).Name for i in range(nb.Elements.Count))))
for tgt in [r'D:\<化工工作区>\NA-Chemical-10000t_最终定稿.apwz']:
    try:
        doc.SaveAs(tgt)
        print('已另存', tgt)
    except Exception as ex:
        print('另存失败', ex)
try:
    doc.Close()
except Exception:
    pass
print('DONE')
