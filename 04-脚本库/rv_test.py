# -*- coding: utf-8 -*-
"""用可靠判据（保存文件中 NRTL 段的 UVAL 条数）重测各写法"""
import os, re, json, time
import win32com.client as w32

BASE = r'D:\<化工工作区>\NA_struct.bkp'
OUT = r'D:\<化工工作区>\_probe'
t = open(BASE, encoding='utf-8', errors='ignore').read()

# 补 TYPE
for outname in ['NAM', 'NAC', '"4-MP"']:
    t, _ = re.subn(r'(OUTNAME = ' + re.escape(outname) + r')(\s+)(DBNAME1)',
                   r'\1 TYPE = CONV\2\3', t, count=1)

m = re.search(r'(TUNITLABEL = C\s*)BDBANK\s*=\s*\([^)]*\)\s*NEL\s*=\s*\d+\s*ESTIMATE\s*=\s*NO', t)
head = m.group(0)
pre = m.group(1)
print('head ok:', bool(head))

# 删掉那条空壳 BPVAL 记录
shell = re.search(r'\\ BPVAL PARAMNAME2 = NRTL CID1 = \s*H2O CID2 = TOL.*?VAL12 = "[^"]*"\s*\\', t, re.S)
print('shell found:', bool(shell))

VAR = {}
VAR['A_base'] = t
VAR['B_nobank'] = t.replace(head, pre + 'NEL = 12 ESTIMATE = NO', 1)
VAR['C_lle'] = t.replace(head, pre + 'BDBANK = ( "APV150 LLE-ASPEN" ) NEL = 1 ESTIMATE = NO', 1)
VAR['D_est'] = t.replace(head, pre + 'NEL = 12 ESTIMATE = YES', 1)
if shell:
    t2 = t[:shell.start()] + '\\' + t[shell.end():]
    VAR['E_noshell'] = t2
    VAR['F_noshell_nobank'] = t2.replace(head, pre + 'NEL = 12 ESTIMATE = NO', 1)

# 加 FILE-SYM-NAM
VAR['G_dbbank'] = VAR['B_nobank'].replace(
    '\\ DATABANKS \\',
    '\\ DATABANKS FILE-SYM-NAM = ( "APV150 PURE41" "APV150 AQUEOUS" "APV150 SOLIDS" "APV150 INORGANIC" "APV150 ASPENPCD" ) \\', 1)


def uval_count(p):
    try:
        s = open(p, encoding='utf-8', errors='ignore').read()
    except Exception:
        return None
    i = s.find('PARAMNAME = NRTL')
    if i < 0:
        return None
    j = s.find('PARAMNAME =', i + 20)
    seg = s[i:j] if j > i else s[i:i+40000]
    return seg.count('UVAL'), seg.count('BPVAL'), len(s)


print()
print('%-20s %8s %6s %8s' % ('variant', 'UVAL', 'BPVAL', 'savedKB'))
res = {}
for tag, txt in VAR.items():
    p = os.path.join(OUT, 'rv_%s.bkp' % tag)
    open(p, 'w', encoding='utf-8', errors='ignore').write(txt)
    sp = os.path.join(OUT, 'rv_%s_s.bkp' % tag)
    try:
        app = w32.DispatchEx('Apwn.Document')
        app.InitFromArchive2(p)
        app.Engine.Run2(False)
        time.sleep(4)
        try:
            app.SaveAs(sp)
        except Exception as e:
            print(tag, 'save ERR', str(e)[:60])
        try:
            app.Close()
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass
    except Exception as e:
        print(tag, 'ERR', str(e)[:80])
        continue
    c = uval_count(sp)
    res[tag] = c
    if c:
        print('%-20s %8d %6d %8d' % (tag, c[0], c[1], c[2] // 1024))
    else:
        print('%-20s   n/a' % tag)

open(os.path.join(OUT, 'rv.json'), 'w', encoding='utf-8').write(json.dumps(res, ensure_ascii=False, indent=1))
print('DONE')
