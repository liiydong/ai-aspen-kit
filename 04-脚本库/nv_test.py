# -*- coding: utf-8 -*-
"""测试 NRTL 参数集头部写法对检索的影响"""
import os, re, json, time
import win32com.client as w32

BASE = r'D:\<化工工作区>\NA_struct.bkp'
OUT = r'D:\<化工工作区>\_probe'
t = open(BASE, encoding='utf-8', errors='ignore').read()

# 先补 TYPE
for outname in ['NAM', 'NAC', '"4-MP"']:
    t, _ = re.subn(r'(OUTNAME = ' + re.escape(outname) + r')(\s+)(DBNAME1)',
                   r'\1 TYPE = CONV\2\3', t, count=1)

m = re.search(r'(TUNITLABEL = C\s*)BDBANK\s*=\s*\([^)]*\)\s*NEL\s*=\s*\d+\s*ESTIMATE\s*=\s*NO', t)
print('NRTL 头部定位:', bool(m))
if m:
    print(repr(m.group(0)[:200]))

VAR = {'base': t}
if m:
    head = m.group(0)
    pre = m.group(1)
    VAR['nel12'] = t.replace(head, pre + 'NEL = 12 ESTIMATE = NO', 1)
    VAR['nobank'] = t.replace(head, pre + 'NEL = 12 ESTIMATE = NO', 1)
    VAR['lleonly'] = t.replace(head, pre + 'BDBANK = ( "APV150 LLE-ASPEN" ) NEL = 1 ESTIMATE = YES', 1)
    VAR['estyes'] = t.replace(head, head.replace('ESTIMATE = NO', 'ESTIMATE = YES'), 1)

res = {}
for tag, txt in VAR.items():
    p = os.path.join(OUT, 'nv_%s.bkp' % tag)
    open(p, 'w', encoding='utf-8', errors='ignore').write(txt)
    rec = {}
    app = w32.DispatchEx('Apwn.Document')
    try:
        app.InitFromArchive2(p)
        try:
            n = app.Tree.FindNode(r'\Data\Components\Specifications\Input\CID')
            rec['comps'] = n.Elements.Count
        except Exception as e:
            rec['comps'] = 'ERR'
        app.Engine.Run2(False)
        time.sleep(3)
        for path in [r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Input\CID1',
                     r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Output\CID1',
                     r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Output\VALUE']:
            try:
                n = app.Tree.FindNode(path)
                rec[path.split('\\')[-1]] = n.Elements.Count if n is not None else None
            except Exception:
                rec[path.split('\\')[-1]] = 'ERR'
    except Exception as e:
        rec['error'] = str(e)[:100]
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

open(os.path.join(OUT, 'nv.json'), 'w', encoding='utf-8').write(json.dumps(res, ensure_ascii=False, indent=1))
print('DONE')
