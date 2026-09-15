# -*- coding: utf-8 -*-
"""对照实验：出厂示例 glycols.bkp 的 NRTL 参数在 COM 里能否读到"""
import os, re, time
import win32com.client as w32

CAND = {
    'glycols': r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\glycols.bkp',
}
GLOBS = []
import glob
GLOBS += glob.glob(r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\**\glycols.bkp', recursive=True)
for g in GLOBS:
    CAND[os.path.basename(g)] = g
print('候选:', CAND)

for tag, p in CAND.items():
    if not os.path.exists(p):
        print(tag, 'missing', p); continue
    app = w32.DispatchEx('Apwn.Document')
    rec = {}
    try:
        app.InitFromArchive2(p)
        print(tag, 'loaded; \Data children =', app.Tree.FindNode(r'\Data').Elements.Count)
        for path in [r'\Data\Properties\Parameters\Binary Interaction',
                     r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Input\CID1',
                     r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Output\CID1',
                     r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Output\VALUE']:
            try:
                n = app.Tree.FindNode(path)
                if n is None:
                    print('   %-64s None' % path); continue
                try:
                    print('   %-64s rows=%d' % (path, n.Elements.Count))
                except Exception:
                    print('   %-64s (leaf)' % path)
            except Exception as e:
                print('   %-64s ERR %s' % (path, str(e)[:40]))
        # 参数表下有哪些子集
        try:
            n = app.Tree.FindNode(r'\Data\Properties\Parameters\Binary Interaction')
            for i in range(n.Elements.Count):
                e = n.Elements.Item(i)
                try:
                    c = e.Elements.Count
                except Exception:
                    c = '?'
                print('      set %-14s children=%s' % (e.Name, c))
        except Exception as e:
            print('      setlist ERR', str(e)[:60])
    except Exception as e:
        print(tag, 'ERROR', str(e)[:120])
    try:
        app.Close()
    except Exception:
        pass
    try:
        app.Quit()
    except Exception:
        pass
print('DONE')
