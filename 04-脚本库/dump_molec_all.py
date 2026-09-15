# -*- coding: utf-8 -*-
"""列出用户文件与完整模型中所有 MOLEC-STRUCT 出现位置的上下文"""
import sys, re, os
sys.stdout.reconfigure(encoding='utf-8')

for P, tag in [(r'D:\<化工工作区>\NA_user.bkp', '用户回存'),
               (r'D:\<化工工作区>\NA_work.bkp', '完整模型')]:
    if not os.path.exists(P):
        print('缺失', P); continue
    x = open(P, encoding='utf-8', errors='ignore').read()
    print('=' * 74)
    print('%s | %d bytes | MOLEC-STRUCT 出现 %d 次' % (tag, len(x), x.count('MOLEC-STRUCT')))
    print('=' * 74)
    for i, m in enumerate(re.finditer(r'MOLEC-STRUCT', x)):
        s = m.start()
        ctx = x[max(0, s - 30): s + 90].replace('\n', '\\n')
        print('  #%-2d @%-7d ...%s...' % (i + 1, s, ctx))
    print()
