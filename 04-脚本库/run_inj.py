# -*- coding: utf-8 -*-
"""跑通注入了真实二元参数的模型，采集关键结果"""
import os, re, time
import win32com.client as w32

P = r'D:\<化工工作区>\_probe\inj.bkp'
OUT = r'D:\<化工工作区>\_probe'

app = w32.DispatchEx('Apwn.Document')
app.InitFromArchive2(P)
print('\Data children =', app.Tree.FindNode(r'\Data').Elements.Count)

app.Engine.Run2(False)
time.sleep(5)

# 状态
try:
    st = app.Tree.FindNode(r'\Data\Results Summary\Run-Status\Output\UOSSTAT2')
    print('UOSSTAT2:', st.Value)
except Exception as e:
    print('status err', str(e)[:60])
for k in ['TERROR', 'SERROR', 'ERROR', 'WARNING']:
    try:
        st = app.Tree.FindNode(r'\Data\Results Summary\Run-Status\Output\%s' % k)
        print('  %-8s = %s' % (k, st.Value))
    except Exception:
        pass


def gv(path):
    try:
        n = app.Tree.FindNode(path)
        return n.Value if n is not None else None
    except Exception:
        return None


STREAMS = ['S-106', 'S-109', 'S-110', 'S-114', 'S-117', 'S-121', 'S-209', 'S-405B', 'S-403']
print()
print('--- 关键流股 ---')
for s in STREAMS:
    m = gv(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s)
    t = gv(r'\Data\Streams\%s\Output\TEMP_OUT\MIXED' % s)
    print('  %-8s MASS=%-12s T=%s' % (s, m, t))

print()
print('--- 物性方法 ---')
for p2 in [r'\Data\Properties\Methods\Input\OPTION_SET',
           r'\Data\Properties\Methods\Selected Methods\Input\OPTION-SETS']:
    print('  %-58s %s' % (p2, gv(p2)))

# NRTL 参数条数（内存）
for p2 in [r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Input\BPVAL',
           r'\Data\Properties\Parameters\Binary Interaction\NRTL-1\Input\CID1']:
    try:
        n = app.Tree.FindNode(p2)
        print('  %-58s %s' % (p2, n.Elements.Count if n is not None else None))
    except Exception as e:
        print('  %-58s ERR' % p2)

# 保存
sp = r'D:\<化工工作区>\NA_withBIP.bkp'
try:
    app.SaveAs(sp)
    print('saved', sp, os.path.getsize(sp))
except Exception as e:
    print('save err', str(e)[:80])
try:
    app.Close()
except Exception:
    pass
try:
    app.Quit()
except Exception:
    pass
print('DONE')
