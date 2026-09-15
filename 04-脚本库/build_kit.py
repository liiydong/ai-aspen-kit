# -*- coding: utf-8 -*-
"""构建 AI-Aspen 封装包"""
import os, shutil, traceback

ROOT = r'D:\obsidian-vault\ai-aspen-kit'
SKILLS = r'<技能根>\skills'
RT = r'D:\<化工工作区>\chemical-engineering-runtime'
CR = r'D:\<化工运行时>'
CASE = r'C:\Users\<用户名>\<项目原始资料>化学法版'
LOGF = r'D:\<化工工作区>\_probe\buildkit.log'
log = []

TEXT_EXT = {'.md', '.py', '.json', '.yaml', '.yml', '.txt', '.csv', '.toml', '.cfg', '.ini', '.html', '.svg'}


def mk(*p):
    d = os.path.join(ROOT, *p)
    os.makedirs(d, exist_ok=True)
    return d


def copy_tree(src, dst, exts=None, skip=('__pycache__', '.git', 'node_modules', '.venv')):
    n = 0
    size = 0
    for base, dirs, files in os.walk(src):
        dirs[:] = [d for d in dirs if d not in skip]
        rel = os.path.relpath(base, src)
        target = os.path.join(dst, rel) if rel != '.' else dst
        for f in files:
            if exts and os.path.splitext(f)[1].lower() not in exts:
                continue
            s = os.path.join(base, f)
            os.makedirs(target, exist_ok=True)
            shutil.copy2(s, os.path.join(target, f))
            n += 1
            size += os.path.getsize(s)
    return n, size


def main():
    # 1. 技能包
    want = ('aspen-', 'chemical-')
    extra = ('equipment-design-app', 'sw6-scripted-equipment-design')
    skills = [d for d in os.listdir(SKILLS)
              if os.path.isdir(os.path.join(SKILLS, d)) and (d.startswith(want) or d in extra)]
    dst = mk('02-技能包')
    for s in sorted(skills):
        n, sz = copy_tree(os.path.join(SKILLS, s), os.path.join(dst, s))
        log.append('SKILL  %-44s %3d files %7.1f KB' % (s, n, sz / 1024))

    # 2. 运行时（仅文本）
    n, sz = copy_tree(RT, mk('03-运行时-文本'), exts=TEXT_EXT)
    log.append('RUNTIME(text) %d files %.1f MB' % (n, sz / 1048576))

    # 3. 探索脚本库
    n, sz = copy_tree(CR, mk('04-脚本库'), exts={'.py'})
    log.append('SCRIPTS %d files %.1f MB' % (n, sz / 1048576))

    # 4. 案例
    d = mk('05-案例-烟酰胺10000ta')
    for f in sorted(os.listdir(CASE)):
        if f.startswith('_'):
            continue
        shutil.copy2(os.path.join(CASE, f), os.path.join(d, f))
    shutil.copy2(os.path.join(CASE, '_数据核对与订正说明.md'),
                 os.path.join(d, '数据核对与订正说明.md'))
    log.append('CASE %d files' % len(os.listdir(d)))

    # 5. 原始资源
    d2 = mk('01-资源与来源')
    for src in [r'D:\<化工工作区>\LOCAL_KNOWLEDGE_GRAPH_LINKS.md',
                r'D:\<化工工作区>\CHEMICAL_SKILLS_INSTALLATION.json',
                r'D:\<化工工作区>\AGENTS.md']:
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(d2, os.path.basename(src)))
            log.append('SOURCE %s' % os.path.basename(src))


try:
    main()
except Exception:
    log.append('ERROR\n' + traceback.format_exc())

tot = 0
cnt = 0
for base, dirs, files in os.walk(ROOT):
    for f in files:
        tot += os.path.getsize(os.path.join(base, f))
        cnt += 1
log.append('')
log.append('TOTAL %d files, %.1f MB' % (cnt, tot / 1048576))
open(LOGF, 'w', encoding='utf-8').write('\n'.join(log))
