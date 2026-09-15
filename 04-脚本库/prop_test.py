# -*- coding: utf-8 -*-
"""测试：以 PROP（仅物性）模式运行，是否触发二元参数检索并写入文件"""
import os, re, shutil, time
import win32com.client as w32

OUT = r'D:\<化工工作区>\_probe'
SRC = r'C:\Users\<用户名>\Desktop\烟酰胺Aspen模型_请补参数后另存.bkp'

cands = {}
if os.path.exists(SRC):
    dst = os.path.join(OUT, 'userprop.bkp')
    shutil.copy(SRC, dst)
    cands['user_prop'] = dst

# 我们的完整模型转成 PROP 模式
t = open(r'D:\<化工工作区>\NA_struct.bkp', encoding='utf-8', errors='ignore').read()
t2, n = re.subn(r'RUN-CLASS\s*=\s*FLOWSHEET', 'RUN-CLASS = PROP', t, count=1)
print('RUN-CLASS 改写次数:', n)
p2 = os.path.join(OUT, 'na_prop.bkp')
open(p2, 'w', encoding='utf-8', errors='ignore').write(t2)
cands['na_as_prop'] = p2

# na_prop + ESTIMATE = YES
t3 = t2.replace('ESTIMATE = NO', 'ESTIMATE = YES')
p3 = os.path.join(OUT, 'na_prop_est.bkp')
open(p3, 'w', encoding='utf-8', errors='ignore').write(t3)
cands['na_prop_est'] = p3


def uval(p):
    if not os.path.exists(p):
        return None
    s = open(p, encoding='utf-8', errors='ignore').read()
    i = s.find('PARAMNAME = NRTL')
    if i < 0:
        return None
    j = s.find('PARAMNAME =', i + 20)
    seg = s[i:j] if j > i else s[i:i+40000]
    return seg.count('UVAL'), seg.count('BPVAL'), len(s), s.count('FILE-SYM-NAM')


for tag, p in cands.items():
    sp = os.path.join(OUT, 'prop_%s_s.bkp' % tag)
    print()
    print('=====', tag, os.path.basename(p))
    try:
        app = w32.DispatchEx('Apwn.Document')
        app.InitFromArchive2(p)
        print('  \Data children =', app.Tree.FindNode(r'\Data').Elements.Count)
        try:
            rc = app.Tree.FindNode(r'\Data\Setup\Input\RUN-CLASS')
            print('  RUN-CLASS =', rc.Value if rc else None)
        except Exception:
            pass
        app.Engine.Run2(False)
        time.sleep(5)
        app.SaveAs(sp)
        print('  saved ->', os.path.basename(sp))
        try:
            app.Close()
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass
    except Exception as e:
        print('  ERR', str(e)[:110])
        continue
    c = uval(sp)
    print('  UVAL=%s BPVAL=%s size=%s SYM=%s' % c if c else '  (no NRTL seg)')
print('DONE')
