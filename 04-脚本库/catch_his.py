# -*- coding: utf-8 -*-
"""切换工作目录后运行，截获 .his 控制面板"""
import os, time, glob
DIR = r'D:\<化工工作区>\_run'
os.makedirs(DIR, exist_ok=True)
for f in glob.glob(os.path.join(DIR, '*')):
    try:
        os.remove(f)
    except Exception:
        pass
os.chdir(DIR)
print('工作目录 =', os.getcwd(), flush=True)

import win32com.client as win32
SRC = r'D:\<化工工作区>\_probe\noair2.bkp'
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(SRC)
time.sleep(4)
t0 = time.time()
try:
    doc.Engine.Run2()
except Exception as e:
    print('Run2 异常', str(e)[:120], flush=True)
while time.time() - t0 < 120:
    time.sleep(3)
    try:
        if doc.Engine.IsRunning is False:
            break
    except Exception:
        break
print('运行用时 %.0fs' % (time.time() - t0), flush=True)
try:
    doc.Close()
except Exception:
    pass

print()
print('=== 工作目录内容 ===', flush=True)
for f in sorted(glob.glob(os.path.join(DIR, '*'))):
    print('   %-40s %d' % (os.path.basename(f), os.path.getsize(f)), flush=True)

print()
for pat in ('*.his', '*.sum', '*.rep', '*.ith', '*.its'):
    for f in glob.glob(os.path.join(DIR, pat)):
        print('=== %s ===（前 5000 字）' % os.path.basename(f), flush=True)
        txt = open(f, encoding='utf-8', errors='ignore').read()
        print(txt[:5000], flush=True)
        print()
