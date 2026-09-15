# -*- coding: utf-8 -*-
"""查 Flash2 端口语义 + 各块端口清单"""
import sys, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32

doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
doc.InitFromArchive2(r'D:\<化工工作区>\NA-Chemical-10000t_s5.bkp')
time.sleep(3)

L = []
for bid in ['E-601', 'E-602', 'C-601', 'D-601', 'T-301', 'M-601', 'T-401', 'T-302', 'M-502']:
    L.append('=========== %s' % bid)
    for p in [r'\Data\Blocks\%s\Ports' % bid, r'\Data\Blocks\%s\Ports\%s' % (bid, bid)]:
        n = doc.Tree.FindNode(p)
        if n is None:
            L.append('  (无) %s' % p)
            continue
        try:
            cnt = n.Elements.Count
        except Exception:
            cnt = None
        L.append('  %s  Elements=%s' % (p, cnt))
        if cnt:
            for i in range(cnt):
                try:
                    e = n.Elements.Item(i)
                    ev = ''
                    try:
                        ev = str(e.Value)
                    except Exception:
                        pass
                    L.append('     %s = %s' % (e.Name, ev))
                except Exception as ex:
                    L.append('     读取失败 %s' % ex)
    L.append('')

open(r'D:\<化工工作区>\_probe\ports.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('DONE')
try:
    doc.Close()
except Exception:
    pass
