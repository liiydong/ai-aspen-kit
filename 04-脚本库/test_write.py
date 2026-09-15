# -*- coding: utf-8 -*-
"""验证 COM 能否写入流股与模块参数"""
import time
import win32com.client as win32

doc = win32.DispatchEx('Apwn.Document')
doc.InitFromArchive2(r'D:\<化工工作区>\_probe\na_from_apwz.bkp')
time.sleep(3)

def rd(p):
    try:
        return doc.Tree.FindNode(p).Value
    except Exception as e:
        return 'ERR:' + str(e)[:60]

def wr(p, v):
    try:
        doc.Tree.FindNode(p).Value = v
        return 'OK'
    except Exception as e:
        return 'ERR:' + str(e)[:90]

print('--- 流股写入测试 ---')
print('S-101 TEMP 原值:', rd(r'\Data\Streams\S-101\Input\TEMP'))
print('  写入 25:', wr(r'\Data\Streams\S-101\Input\TEMP', 25.0))
print('S-101 TEMP 现在:', rd(r'\Data\Streams\S-101\Input\TEMP'))
print('S-101 PRES 原值:', rd(r'\Data\Streams\S-101\Input\PRES'))
print('  写入 2.0:', wr(r'\Data\Streams\S-101\Input\PRES', 2.0))
print('S-101 PRES 现在:', rd(r'\Data\Streams\S-101\Input\PRES'))

print('\n--- 模块写入测试 ---')
print('R-101 TEMP 原值:', rd(r'\Data\Blocks\R-101\Input\TEMP'))
print('  写入 405:', wr(r'\Data\Blocks\R-101\Input\TEMP', 405.0))
print('R-101 TEMP 现在:', rd(r'\Data\Blocks\R-101\Input\TEMP'))
print('R-101 PRES 原值:', rd(r'\Data\Blocks\R-101\Input\PRES'))
print('  写入 1.8:', wr(r'\Data\Blocks\R-101\Input\PRES', 1.8))
print('R-101 PRES 现在:', rd(r'\Data\Blocks\R-101\Input\PRES'))

print('\n--- 检查流股的关键参数字段是否存在 ---')
for f in ('TEMP', 'PRES', 'MASS-FLOW', 'MOLE-FLOW', 'MASSFRAC', 'MOLEFRAC', 'VFRAC'):
    print(f'  {f}:', rd(r'\Data\Streams\S-101\Input\\' + f))
