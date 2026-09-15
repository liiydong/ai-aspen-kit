# -*- coding: utf-8 -*-
"""把 Aspen xlsx 报告按 A/B/C 三列原样导出（TAB 分隔），供解析脚本重建固定宽度行。"""
import openpyxl, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

SRC = r'C:\Users\<用户名>\Desktop\ASPEN PLUS-最终版报告-2026.xlsx'
OUT = r'D:\<化工工作区>\_probe\report_cols.tsv'

wb = openpyxl.load_workbook(SRC, data_only=True)
ws = wb['Sheet1']
lines = []
for r in range(1, ws.max_row + 1):
    vals = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=r, column=c).value
        vals.append('' if v is None else str(v))
    # 去掉纯空行
    if all(x.strip() == '' for x in vals):
        continue
    lines.append('%d\t%s' % (r, '\t'.join(vals)))

with open(OUT, 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))
print('rows:', len(lines))
