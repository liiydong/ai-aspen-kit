# -*- coding: utf-8 -*-
"""一次性测试：声明 DATABANKS 库清单后，Aspen 能否自动检索到 NRTL 二元参数数值"""
import os, re, json, time
import win32com.client as w32

BASE = r'D:\<化工工作区>\NA_struct.bkp'
OUTDIR = r'D:\<化工工作区>\_probe'

t = open(BASE, encoding='utf-8', errors='ignore').read()
NEEDLE = '\\ DATABANKS \\'
print('锚点出现次数:', t.count(NEEDLE))

VARIANTS = {
    'none': None,
    'fac4': '\\ DATABANKS FILE-SYM-NAM = ( "APV150 PURE41" "APV150 AQUEOUS" "APV150 INORGANIC" "APV150 ASPENPCD" ) \\',
    'pure4': '\\ DATABANKS FILE-SYM-NAM = ( "APV150 PURE41" "APV150 AQUEOUS" "APV150 SOLIDS" "APV150 INORGANIC" ) \\',
    'pure5': '\\ DATABANKS FILE-SYM-NAM = ( "APV150 PURE41" "APV150 AQUEOUS" "APV150 SOLIDS" "APV150 INORGANIC" "APV150 ASPENPCD" ) \\',
}

def make(tag, text):
    p = os.path.join(OUTDIR, 'db_%s.bkp' % tag)
    open(p, 'w', encoding='utf-8', errors='ignore').write(text)
    return p

results = {}
for tag, rep in VARIANTS.items():
    txt = t if rep is None else t.replace(NEEDLE, rep, 1)
    p = make(tag, txt)
    rec = {'file': p}
    try:
        app = w32.DispatchEx('Apwn.Document')
        try:
            app.InitFromArchive2(p)
        except Exception as e:
            rec['load'] = 'FAIL ' + str(e)[:80]
            app.Quit(); continue
        # 读 DATABANKS / 参数表
        try:
            fsn = app.Tree.FindNode(r'\Data\Databanks\Input\FILE_SYM_NAM')
            rec['FILE_SYM_NAM'] = fsn.Elements.Count if fsn is not None else None
        except Exception as e:
            rec['FILE_SYM_NAM'] = 'ERR ' + str(e)[:60]
        app.Engine.Run2(False)
        time.sleep(2)
        try:
            rec['msgs'] = app.Engine.Messages.Count if hasattr(app.Engine, 'Messages') else None
        except Exception:
            rec['msgs'] = None
        for path in [r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Output\CID1',
                     r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Input\VAL1']:
            try:
                n = app.Tree.FindNode(path)
                rec[path.split('\\')[-1]] = n.Elements.Count if n is not None else None
            except Exception as e:
                rec[path.split('\\')[-1]] = 'ERR'
        # 保存后检查 DATABANKS 是否保留
        sp = os.path.join(OUTDIR, 'db_%s_s.bkp' % tag)
        try:
            app.SaveAs(sp)
            s = open(sp, encoding='utf-8', errors='ignore').read()
            rec['saved_syms'] = 'FILE-SYM-NAM' in s
            rec['saved_size'] = len(s)
        except Exception as e:
            rec['saved'] = 'ERR ' + str(e)[:80]
        try:
            app.Close()
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass
    except Exception as e:
        rec['error'] = str(e)[:120]
    results[tag] = rec
    print(tag, json.dumps(rec, ensure_ascii=False)[:400])

open(os.path.join(OUTDIR, 'db_test.json'), 'w', encoding='utf-8').write(json.dumps(results, ensure_ascii=False, indent=1))
print('DONE')
