# -*- coding: utf-8 -*-
"""修复 04-脚本库（仅顶层脚本）并补齐 05-案例"""
import os, shutil, traceback

ROOT = r'D:\obsidian-vault\ai-aspen-kit'
CR = r'D:\<化工运行时>'
CASE = r'C:\Users\<用户名>\<项目原始资料>化学法版'
LOGF = r'D:\<化工工作区>\_probe\fix_kit.log'
log = []

try:
    # --- 1. 重建脚本库：只取顶层 .py ---
    dst = os.path.join(ROOT, '04-脚本库')
    readme = os.path.join(dst, 'README.md')
    tmp = None
    if os.path.exists(readme):
        tmp = os.path.join(ROOT, '_readme_scripts.bak')
        shutil.move(readme, tmp)
    shutil.rmtree(dst, ignore_errors=True)
    os.makedirs(dst, exist_ok=True)
    n = sz = 0
    for f in sorted(os.listdir(CR)):
        p = os.path.join(CR, f)
        if os.path.isfile(p) and f.lower().endswith('.py'):
            try:
                shutil.copy2(p, os.path.join(dst, f))
                n += 1
                sz += os.path.getsize(p)
            except Exception as e:
                log.append('  skip %s (%s)' % (f, e.__class__.__name__))
    log.append('SCRIPTS(top-level) %d files %.1f MB' % (n, sz / 1048576))
    if tmp and os.path.exists(tmp):
        shutil.move(tmp, readme)

    # --- 2. 案例 ---
    d = os.path.join(ROOT, '05-案例-烟酰胺10000ta')
    os.makedirs(d, exist_ok=True)
    for f in sorted(os.listdir(CASE)):
        if f.startswith('_'):
            continue
        shutil.copy2(os.path.join(CASE, f), os.path.join(d, f))
    shutil.copy2(os.path.join(CASE, '_数据核对与订正说明.md'),
                 os.path.join(d, '数据核对与订正说明.md'))
    log.append('CASE %d files' % len(os.listdir(d)))
except Exception:
    log.append('ERROR\n' + traceback.format_exc())

tot = cnt = 0
for base, dirs, files in os.walk(ROOT):
    for f in files:
        tot += os.path.getsize(os.path.join(base, f))
        cnt += 1
log.append('TOTAL %d files, %.1f MB' % (cnt, tot / 1048576))
open(LOGF, 'w', encoding='utf-8').write('\n'.join(log))
