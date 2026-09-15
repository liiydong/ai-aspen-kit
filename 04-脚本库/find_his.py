# -*- coding: utf-8 -*-
"""全盘搜索 Aspen 运行产生的 .his/.sum/.rep，并 dump EngineFilesSettings"""
import os, time, glob
import win32com.client as win32

print('=== EngineFilesSettings / OptionSettings ===', flush=True)
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(r'D:\<化工工作区>\_probe\noair2.bkp')
time.sleep(4)
for obj in ['EngineFilesSettings', 'OptionSettings', 'RunSettings']:
    o = getattr(doc.Engine, obj, None)
    if o is None:
        print('  %s 无' % obj, flush=True)
        continue
    try:
        names = [m for m in dir(o) if not m.startswith('_')][:20]
        print('  %s 成员: %s' % (obj, names), flush=True)
        for m in names:
            try:
                v = getattr(o, m)
                if not callable(v) and v is not None:
                    print('     %s = %r' % (m, v), flush=True)
            except Exception:
                pass
    except Exception as e:
        print('  %s 读取失败 %s' % (obj, str(e)[:80]), flush=True)
try:
    doc.Close()
except Exception:
    pass

print()
print('=== 搜索最近的 Aspen 输出文件 ===', flush=True)
roots = [r'D:\<化工工作区>', r'C:\Users\<用户名>\AppData\Local\Temp', r'C:\Users\<用户名>\Documents',
         r'D:\<化工运行时>', r'C:\Users\<用户名>']
found = []
now = time.time()
for r in roots:
    for ext in ('his', 'sum', 'rep', 'ith', 'its'):
        try:
            for p in glob.glob(os.path.join(r, '**', '*.' + ext), recursive=True):
                try:
                    mt = os.path.getmtime(p)
                    if now - mt < 6 * 3600:
                        found.append((mt, p, os.path.getsize(p)))
                except Exception:
                    pass
        except Exception:
            pass
found.sort(reverse=True)
print('  找到 %d 个（近 6 小时）' % len(found), flush=True)
for mt, p, sz in found[:20]:
    print('   %s  %8d  %s' % (time.strftime('%H:%M:%S', time.localtime(mt)), sz, p), flush=True)
