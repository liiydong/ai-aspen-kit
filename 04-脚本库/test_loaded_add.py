# -*- coding: utf-8 -*-
"""在已加载的真实模型上测试：新建模块 / 删除模块"""
import time
import win32com.client as win32

REF = r'D:\<化工工作区>\_probe\ref_multi.bkp'
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(REF)
time.sleep(3)

blk = doc.Tree.FindNode(r'\Data\Blocks')
print('现有模块数:', blk.Elements.Count)
print('模块名:', [blk.Elements.Item(i).Name for i in range(blk.Elements.Count)])
print()

CANDS = [
    ('Add("ZZ1")', ('ZZ1',)),
    ('Add("ZZ1","Mixer")', ('ZZ1', 'Mixer')),
    ('Add("ZZ1","MIXER")', ('ZZ1', 'MIXER')),
    ('Add("ZZ1","Mixer","Built-In")', ('ZZ1', 'Mixer', 'Built-In')),
    ('Add("ZZ2","RADFRAC")', ('ZZ2', 'RADFRAC')),
    ('Add("ZZ3","Heater")', ('ZZ3', 'Heater')),
]
for desc, args in CANDS:
    try:
        blk.Elements.Add(*args)
        print('  %-36s 成功!  模块数=%s' % (desc, blk.Elements.Count))
    except Exception as e:
        print('  %-36s 失败 %s' % (desc, str(e)[:75]))
print()

st = doc.Tree.FindNode(r'\Data\Streams')
print('现有流股数:', st.Elements.Count)
print('流股名:', [st.Elements.Item(i).Name for i in range(st.Elements.Count)])
try:
    st.Elements.Add('ZZS1')
    print('  新增流股成功!  流股数=%s' % st.Elements.Count)
except Exception as e:
    print('  新增流股失败', str(e)[:80])
print()

# 删除测试（删掉刚加的）
for nm in ['ZZS1']:
    try:
        st.Elements.Remove(nm)
        print('  删除 %s 成功, 剩 %s' % (nm, st.Elements.Count))
    except Exception as e:
        print('  删除 %s 失败 %s' % (nm, str(e)[:70]))

try:
    doc.Close()
except Exception:
    pass
