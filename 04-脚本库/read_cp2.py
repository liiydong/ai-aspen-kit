# -*- coding: utf-8 -*-
"""运行后读控制面板 / 导出报告"""
import time, os
import win32com.client as win32

F = r'D:\<化工工作区>\NA-Chemical-10000t_模拟.bkp'
RPT = r'D:\<化工工作区>\_probe\cp_report.txt'
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(F)
time.sleep(4)

print('=== 运行 ===', flush=True)
try:
    doc.Engine.Run2()
    print('  Run2 完成', flush=True)
except Exception as e:
    print('  Run2 抛出:', str(e)[:150], flush=True)
time.sleep(3)

print()
print('=== Engine.ControlPanel（运行后）===', flush=True)
try:
    cp = doc.Engine.ControlPanel
    print('  类型', type(cp), flush=True)
    try:
        print('  Value:', str(cp.Value)[:3000], flush=True)
    except Exception as e:
        print('  Value 失败', str(e)[:100], flush=True)
except Exception as e:
    print('  失败', str(e)[:130], flush=True)

print()
print('=== Engine.ExportReport ===', flush=True)
try:
    doc.Engine.ExportReport(RPT)
    print('  已导出', RPT, os.path.getsize(RPT), flush=True)
    txt = open(RPT, encoding='utf-8', errors='ignore').read()
    print(txt[:4000], flush=True)
except Exception as e:
    print('  失败', str(e)[:150], flush=True)

print()
print('=== 文件里找 ERROR/WARNING ===', flush=True)
try:
    doc.SaveAs(r'D:\<化工工作区>\_probe\after_run.bkp')
    t = open(r'D:\<化工工作区>\_probe\after_run.bkp', encoding='utf-8', errors='ignore').read()
    import re
    for m in re.finditer(r'.{0,100}(ERROR|WARNING|AE_|\*\*\*).{0,160}', t):
        print('  ', re.sub(r'\s+', ' ', m.group())[:220], flush=True)
except Exception as e:
    print('  失败', str(e)[:120], flush=True)
try:
    doc.Close()
except Exception:
    pass
