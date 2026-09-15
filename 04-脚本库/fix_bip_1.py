# -*- coding: utf-8 -*-
"""取全新模型的 DATABANKS 声明；并检查现有组分的 UNIFAC 基团"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

FRESH = r'D:\<化工工作区>\_probe\fresh_default.bkp'
L = []

# ---------- A) 新建一个空白模型，取默认 DATABANKS ----------
try:
    d = win32.DispatchEx('Apwn.Document')
    d.SuppressDialogs = True
    d.InitNew2()
    time.sleep(2)
    d.SaveAs(FRESH)
    time.sleep(1)
    d.Close()
    L.append('fresh 模型已建: %s (%d bytes)' % (FRESH, os.path.getsize(FRESH)))
except Exception as ex:
    L.append('新建失败: %s' % ex)

if os.path.exists(FRESH):
    t = open(FRESH, encoding='utf-8', errors='ignore').read()
    m = re.search(r'\?\s*DATABANKS\s*\?', t)
    if m:
        j = re.search(r'\n\?\s*[A-Z]', t[m.end():])
        e = m.end() + (j.start() if j else 1200)
        L.append('===== 全新模型的 DATABANKS 段 =====')
        L.append(t[m.start():min(e, m.start() + 1500)])
    else:
        L.append('未找到 DATABANKS 段')
    # 默认物性方法段
    m2 = re.search(r'\?\s*PROPERTIES\s+MAIN\s*\?', t)
    if m2:
        j2 = re.search(r'\n\?\s*[A-Z]', t[m2.end():])
        e2 = m2.end() + (j2.start() if j2 else 600)
        L.append('')
        L.append('===== 全新模型的 PROPERTIES MAIN 段 =====')
        L.append(t[m2.start():min(e2, m2.start() + 700)])

# ---------- B) 现有模型：组分与 UNIFAC 基团 ----------
SRC = r'D:\<化工工作区>\NA-Chemical-10000t_Submit.bkp'
doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
doc.InitFromArchive2(SRC)
time.sleep(2)

def g(p):
    n = doc.Tree.FindNode(p)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return None

L.append('')
L.append('===== 路径探测 =====')
for p in [r'\Data\Properties\Parameters\Binary\T-DEPENDENT\NRTL-1',
          r'\Data\Properties\Parameters\Binary\NRTL-1',
          r'\Data\Properties\Parameters',
          r'\Data\Components\Specifications\Input\CID',
          r'\Data\Components\UNIFAC-Groups',
          r'\Data\Setup\Databanks',
          r'\Data\Setup\Global\Input\DATABANKS',
          r'\Data\Properties\Parameters\Binary\T-DEPENDENT\NRTL-1\Input\BPVAL']:
    n = doc.Tree.FindNode(p)
    L.append('  %-62s %s' % (p, 'None' if n is None else 'OK  children=%s' % (getattr(n, 'Elements', None) and n.Elements.Count)))

# NRTL-1 下的子节点
for p in [r'\Data\Properties\Parameters\Binary\T-DEPENDENT\NRTL-1\Input',
          r'\Data\Properties\Parameters\Binary\T-DEPENDENT\NRTL-1\Output']:
    n = doc.Tree.FindNode(p)
    if n is not None:
        L.append('  ---- %s : %d children' % (p, n.Elements.Count))
        for i in range(min(n.Elements.Count, 30)):
            e = n.Elements.Item(i)
            L.append('       %s = %s' % (e.Name, str(e.Value)[:40]))

# 组分 ID 列表
n = doc.Tree.FindNode(r'\Data\Components\Specifications\Input\CID')
if n is not None:
    cids = []
    for i in range(n.Elements.Count):
        cids.append(str(n.Elements.Item(i).Value))
    L.append('  组分列表: %s' % ', '.join(cids))

# UNIFAC 基团是否有值
n = doc.Tree.FindNode(r'\Data\Components\UNIFAC-Groups')
L.append('  UNIFAC-Groups 节点: %s' % ('None' if n is None else 'children=%d' % n.Elements.Count))

try:
    doc.Close()
except Exception:
    pass

open(r'D:\<化工工作区>\_probe\bip_fix1.txt', 'w', encoding='utf-8').write('\n'.join(L))
print('\n'.join(L))
