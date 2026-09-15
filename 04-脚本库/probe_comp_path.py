# -*- coding: utf-8 -*-
import os, re, json, time
import win32com.client as w32

P = r'D:\<化工工作区>\_probe\ct_fixtype.bkp'
app = w32.DispatchEx('Apwn.Document')
app.InitFromArchive2(P)


def children(path, limit=30):
    try:
        n = app.Tree.FindNode(path)
    except Exception as e:
        return ['ERR ' + str(e)[:60]]
    if n is None:
        return ['None']
    out = []
    try:
        c = n.Elements.Count
    except Exception:
        return ['no Elements: val=' + str(getattr(n, 'Value', '?'))[:60]]
    for i in range(min(c, limit)):
        try:
            e = n.Elements.Item(i)
            out.append('%s [%s]' % (e.Name, str(getattr(e, 'Value', ''))[:24]))
        except Exception as ex:
            out.append('item%d ERR %s' % (i, str(ex)[:40]))
    return out

for path in [r'\Data', r'\Data\Components', r'\Data\Components\Input',
             r'\Data\Properties\Parameters', r'\Data\Properties\Parameters\Binary Interaction',
             r'\Data\Properties\Parameters\Binary Interaction\NRTL-1']:
    print('==', path)
    for c in children(path):
        print('   ', c)
    print()

# 运行并读状态
try:
    app.Engine.Run2(False)
    time.sleep(2)
except Exception as e:
    print('run err', str(e)[:80])

for path in [r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Input\CID1',
             r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Output\CID1',
             r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Input\VAL1']:
    try:
        n = app.Tree.FindNode(path)
        print('%-70s rows=%s' % (path, n.Elements.Count if n is not None else None))
    except Exception as e:
        print('%-70s ERR %s' % (path, str(e)[:50]))

try:
    app.Close()
except Exception:
    pass
try:
    app.Quit()
except Exception:
    pass
print('DONE')
