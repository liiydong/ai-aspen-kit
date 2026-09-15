# -*- coding: utf-8 -*-
"""测试：补上 NAM/NAC/4-MP 的 TYPE = CONV 后，二元参数能否被检索到"""
import os, re, json, time
import win32com.client as w32

BASE = r'D:\<化工工作区>\NA_struct.bkp'
OUT = r'D:\<化工工作区>\_probe'
t = open(BASE, encoding='utf-8', errors='ignore').read()


def add_type(txt):
    n = 0
    for cid, outname in [('NAM', 'NAM'), ('NAC', 'NAC'), ('4-MP', '"4-MP"')]:
        pat = re.compile(r'(OUTNAME = ' + re.escape(outname) + r')(\s+)(DBNAME1)')
        txt, k = pat.subn(r'\1 TYPE = CONV\2\3', txt, count=1)
        n += k
    return txt, n


VAR = {'base': t}
t2, cnt = add_type(t)
print('TYPE 补入次数:', cnt)
VAR['fixtype'] = t2

res = {}
for tag, txt in VAR.items():
    p = os.path.join(OUT, 'ct_%s.bkp' % tag)
    open(p, 'w', encoding='utf-8', errors='ignore').write(txt)
    rec = {}
    app = w32.DispatchEx('Apwn.Document')
    try:
        app.InitFromArchive2(p)
        rec['comp_count'] = app.Tree.FindNode(r'\Data\Components\Input\CID').Elements.Count
        app.Engine.Run2(False)
        time.sleep(2)
        for path in [r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Output\CID1',
                     r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Input\CID1']:
            try:
                n = app.Tree.FindNode(path)
                rec[path.split('\\')[-2] + '/' + path.split('\\')[-1]] = n.Elements.Count if n is not None else None
            except Exception as e:
                rec['err_' + path.split('\\')[-2]] = str(e)[:50]
        # 抓消息里的关键行
        msgs = []
        try:
            mnode = app.Tree.FindNode(r'\Data\Results Summary\Run-Status\Output\UOSSTAT2')
            rec['status'] = mnode.Value if mnode is not None else None
        except Exception:
            pass
    except Exception as e:
        rec['error'] = str(e)[:120]
    try:
        app.Close()
    except Exception:
        pass
    try:
        app.Quit()
    except Exception:
        pass
    res[tag] = rec
    print(tag, json.dumps(rec, ensure_ascii=False))

open(os.path.join(OUT, 'ct.json'), 'w', encoding='utf-8').write(json.dumps(res, ensure_ascii=False, indent=1))
print('DONE')
