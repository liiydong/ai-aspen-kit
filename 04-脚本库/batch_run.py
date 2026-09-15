# -*- coding: utf-8 -*-
"""导出 .inp 并用批处理引擎运行，读取 .his 控制面板"""
import os, sys, time, subprocess, glob

VENV = r'D:\<化工运行时>\python-env\Scripts\python.exe'
WORK = r'D:\<化工工作区>'
ASTOP = r'C:\Program Files\AspenTech\Aspen Plus V15.0'
ENGINE = os.path.join(ASTOP, 'Engine', 'Xeq', 'aspen.exe')
SRC = os.path.join(WORK, '_probe', 'full3.bkp')
INP = os.path.join(WORK, 'NAsim.inp')

print('=== 1) COM 导出 .inp ===', flush=True)
import win32com.client as win32
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(SRC)
time.sleep(4)
print('  Ready =', doc.Engine.Ready, flush=True)
exported = False
for m in ['Export', 'ExportInput', 'WriteInput']:
    fn = getattr(doc, m, None)
    if fn is None:
        print('  %s 不存在' % m, flush=True)
        continue
    try:
        fn(INP)
        print('  %s 成功 -> %s (%d bytes)' % (m, INP, os.path.getsize(INP)), flush=True)
        exported = True
        break
    except Exception as e:
        print('  %s 失败 %s' % (m, str(e)[:100]), flush=True)
try:
    doc.Close()
except Exception:
    pass

if not exported:
    # 退而求其次：直接拿 bkp 当输入
    import shutil
    shutil.copyfile(SRC, INP)
    print('  改用 bkp 副本作为输入', flush=True)

print()
print('=== 2) 批处理运行 ===', flush=True)
env = dict(os.environ)
env['asptop'] = ASTOP
env['ASPROD'] = ASTOP
env['aspen'] = ASTOP
os.chdir(WORK)
for f in glob.glob(os.path.join(WORK, 'NAsim.his')) + glob.glob(os.path.join(WORK, 'NAsim.sum')):
    try:
        os.remove(f)
    except Exception:
        pass

cmd = [ENGINE, 'D:\\<化工工作区>\\NAsim.inp']
print('  cmd:', ' '.join(cmd), flush=True)
try:
    p = subprocess.run(cmd, env=env, cwd=WORK, capture_output=True, timeout=420)
    out = (p.stdout or b'').decode('utf-8', 'ignore') + (p.stderr or b'').decode('utf-8', 'ignore')
    print('  退出码:', p.returncode, flush=True)
    print('  ---- 引擎输出 ----', flush=True)
    print(out[:6000], flush=True)
except subprocess.TimeoutExpired:
    print('  运行超时', flush=True)
except Exception as e:
    print('  运行异常:', str(e)[:200], flush=True)

print()
print('=== 3) 找 .his / .sum / .rep ===', flush=True)
for pat in ['NAsim.his', 'NAsim.sum', 'NAsim.rep', 'NAsim.sta', 'NAsim.cod']:
    p = os.path.join(WORK, pat)
    if os.path.exists(p):
        print('  %s  %d bytes' % (pat, os.path.getsize(p)), flush=True)
    else:
        print('  %s  无' % pat, flush=True)
his = os.path.join(WORK, 'NAsim.his')
if os.path.exists(his):
    txt = open(his, encoding='utf-8', errors='ignore').read()
    print()
    print('=== .his 内容（前 6000 字）===', flush=True)
    print(txt[:6000], flush=True)
