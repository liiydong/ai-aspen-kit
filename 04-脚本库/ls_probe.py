import os, sys, glob, time
sys.stdout.reconfigure(encoding='utf-8')

print('=== _probe 目录文件（按时间） ===')
fs = []
for p in glob.glob(r'D:/<化工工作区>/_probe/*'):
    try:
        fs.append((os.path.getmtime(p), os.path.getsize(p), os.path.basename(p)))
    except Exception:
        pass
for mt, sz, n in sorted(fs)[-25:]:
    print('  %s  %9d  %s' % (time.strftime('%H:%M:%S', time.localtime(mt)), sz, n))

print()
print('=== ChemicalWork 顶层 bkp ===')
for p in sorted(glob.glob(r'D:/<化工工作区>/*.bkp')):
    print('  %s  %9d  %s' % (time.strftime('%H:%M:%S', time.localtime(os.path.getmtime(p))),
                             os.path.getsize(p), os.path.basename(p)))
