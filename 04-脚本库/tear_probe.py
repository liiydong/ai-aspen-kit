# -*- coding: utf-8 -*-
"""看参考模型 TEAR 段 + 给我们的模型补撕裂流 + 再运行"""
import time, re
import win32com.client as win32

REF = r'D:\<化工工作区>\_probe\ref_multi.bkp'
SRC = r'D:\<化工工作区>\_probe\noair2.bkp'
OUT = r'D:\<化工工作区>\_probe\tear1.bkp'

rt = open(REF, encoding='utf-8', errors='ignore').read()
i = rt.find('? TEAR ?')
print('=== 参考模型 TEAR 段 ===', flush=True)
if i >= 0:
    m = re.search(r'\?\s+\S', rt[i + 8:])
    j = i + 8 + m.start() if m else i + 300
    print(repr(rt[i:j][:800]), flush=True)
else:
    print('  没找到 ? TEAR ?', flush=True)
print()

# CONVERGENCE 段
k = rt.find('? CONVERGENCE')
print('=== 参考模型 CONVERGENCE 段 ===', flush=True)
print(repr(rt[k:k + 300]) if k >= 0 else '  无', flush=True)
print()

# 我们的模型
t = open(SRC, encoding='utf-8', errors='ignore').read()
print('=== 我们的模型 ===', flush=True)
print('  有 ? TEAR ?:', '? TEAR ?' in t, flush=True)
k2 = t.find('? CONVERGENCE')
print('  CONVERGENCE 段:', repr(t[k2:k2 + 300]) if k2 >= 0 else '无', flush=True)
m2 = re.search(r'\?\s*SOLVE\s*\?', t)
print('  SOLVE 段:', repr(t[m2.start():m2.start() + 200]) if m2 else '无', flush=True)

# 看我们的收敛树
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(SRC)
time.sleep(4)
cv = doc.Tree.FindNode(r'\Data\Convergence')
print()
print('  我们 Convergence 子节点:', [cv.Elements.Item(i3).Name
      for i3 in range(cv.Elements.Count)], flush=True)
for sub in ['Tear', 'Options', 'Advanced']:
    n = doc.Tree.FindNode(r'\Data\Convergence\%s' % sub)
    if n is None:
        print('    %s: 无' % sub, flush=True)
        continue
    try:
        print('    %s: 子=%d %s' % (sub, n.Elements.Count,
              [n.Elements.Item(i3).Name for i3 in range(min(10, n.Elements.Count))]), flush=True)
    except Exception as e:
        print('    %s: %r' % (sub, n.Value), flush=True)
try:
    doc.Close()
except Exception:
    pass
