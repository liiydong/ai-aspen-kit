# -*- coding: utf-8 -*-
"""创建桌面提交初版文件夹并归档"""
import os, shutil, glob, sys
sys.stdout.reconfigure(encoding='utf-8')
SRC = r'D:\obsidian-vault\毕业设计\化学法版'
CW = r'D:\<化工工作区>'
DST = r'C:\Users\<用户名>\Desktop\烟酰胺毕设_提交初版_20260913'
LOG = []

subs = ['01_论文（提交版）', '02_Aspen模型', '03_附录图纸', '04_数据与报告']
for s in subs:
    os.makedirs(os.path.join(DST, s), exist_ok=True)

def cp(src, d, name=None):
    if not os.path.exists(src):
        LOG.append('  MISS %s' % src); return
    tgt = os.path.join(DST, d, name or os.path.basename(src))
    shutil.copy2(src, tgt)
    LOG.append('  OK   %-46s %8d' % (os.path.join(d, os.path.basename(tgt)), os.path.getsize(tgt)))

LOG.append('=== 01_论文（提交版）===')
for f in sorted(glob.glob(os.path.join(SRC, '0[0-7]_*.docx'))) + sorted(glob.glob(os.path.join(SRC, '0[0-7]_*.md'))):
    cp(f, subs[0])

LOG.append('=== 02_Aspen模型 ===')
for n in ['NA-Chemical-10000t_Submit.bkp', 'NA-Chemical-10000t_Submit.apwz']:
    cp(os.path.join(CW, n), subs[1])

LOG.append('=== 03_附录图纸 ===')
for n in ['C-1_BFD.svg', 'C-2_PFD.svg']:
    cp(os.path.join(SRC, n), subs[2])

LOG.append('=== 04_数据与报告 ===')
for n in ['Aspen流程复核与论文修订清单.md', '提交版检查报告.md', 'Aspen全流程校核报告_定稿.md',
          'Aspen技术评审_瑞邦对齐评估与二元参数攻关.md', 'Aspen_V15操作手册_烟酰胺_化学法_中文界面版.md',
          'Aspen骨架搭建卡_化学法.md', 'Aspen组分输入卡_化学法.md', '化学法版_文档检查报告.md',
          '物料衡算_化学法_烟酰胺10000t.xlsx']:
    cp(os.path.join(SRC, n), subs[3])

# 同时更新旧桌面副本（保持版本一致）
OLD = [r'C:\Users\<用户名>\Desktop\烟酰胺毕设_化学法版_20260910',
       r'C:\Users\<用户名>\<项目原始资料>\_项目归档_20260910\02_论文_化学法版_20260910']
LOG.append('=== 同步旧桌面副本 ===')
for d in OLD:
    if not os.path.isdir(d):
        LOG.append('  MISS %s' % d); continue
    for f in sorted(glob.glob(os.path.join(SRC, '0[0-7]_*.docx'))) + sorted(glob.glob(os.path.join(SRC, '0[0-7]_*.md'))):
        shutil.copy2(f, d)
    for n in ['Aspen流程复核与论文修订清单.md', '提交版检查报告.md']:
        shutil.copy2(os.path.join(SRC, n), d)
    for n in ['NA-Chemical-10000t_Submit.bkp', 'NA-Chemical-10000t_Submit.apwz']:
        shutil.copy2(os.path.join(CW, n), d)
    LOG.append('  OK   同步完成 %s' % d)

# 统计
tot = sum(len(fs) for _, _, fs in os.walk(DST))
LOG.append('')
LOG.append('提交文件夹文件总数（含子目录）: %d' % tot)

open(os.path.join(CW, '_probe', 'build_submit_log.txt'), 'w', encoding='utf-8').write('\n'.join(LOG))
print('\n'.join(LOG))
