# -*- coding: utf-8 -*-
"""对照参考模型 vs 我们的模型：塔的关键字段名与值"""
import time
import win32com.client as win32


def dump(path, tag, blocks):
    print('=' * 16, tag, flush=True)
    doc = win32.DispatchEx('Apwn.Document')
    try:
        doc.SuppressDialogs = True
    except Exception:
        pass
    doc.InitFromArchive(path)
    time.sleep(4)
    print('  Ready =', doc.Engine.Ready, flush=True)
    for b in blocks:
        n = doc.Tree.FindNode(r'\Data\Blocks\%s\Input' % b)
        if n is None:
            print('  %s: 无' % b, flush=True)
            continue
        hits = []
        for i in range(n.Elements.Count):
            ch = n.Elements.Item(i)
            if any(k in ch.Name.upper() for k in
                   ['PRES', 'NSTAGE', 'FEED', 'PROD', 'RR', 'D:F', 'B:F', 'CONDENSER',
                    'REBOILER', 'BASIS']):
                try:
                    v = ch.Value
                except Exception:
                    v = '?'
                if v is not None and v != '':
                    hits.append('%s=%r' % (ch.Name, v))
        print('  [%s] %s' % (b, '; '.join(hits[:14])), flush=True)
    try:
        doc.Close()
    except Exception:
        pass
    print(flush=True)


dump(r'D:\<化工工作区>\_probe\ref_multi.bkp', '参考模型', ['B1', 'B2'])
dump(r'D:\<化工工作区>\NA-Chemical-10000t_模拟.bkp', '我们的模型',
     ['T-401', 'T-201', 'T-301', 'M-101'])
