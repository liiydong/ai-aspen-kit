# -*- coding: utf-8 -*-
"""自己动手：删 AIR -> 试写温压 -> 运行并读控制面板"""
import time, re
import win32com.client as win32

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_工作版.bkp'
OUT = r'D:\<化工工作区>\_probe\noair.bkp'

t = open(SRC, encoding='utf-8', errors='ignore').read()
start = t.find('? COMPONENTS MAIN ?')
end = t.find('\\ ? COMPONENTS "ADA/PCS"', start)
seg = t[start:end]
parts = re.split(r'(\s/\s)', seg)
# 重新拼装，剔除 AIR
kept, removed = [], []
i = 0
while i < len(parts):
    p = parts[i]
    if re.match(r'^\s/\s$', p):
        kept.append(p)
        i += 1
        continue
    if re.search(r'CID = AIR\b', p):
        removed.append(p.strip()[:80])
    else:
        kept.append(p)
    i += 1
newseg = ''.join(kept)
# 清理开头可能残留的分隔符
newseg = re.sub(r'^(\s/\s)+', '', newseg)
t2 = t[:start] + newseg + t[end:]
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t2)
print('已删除 AIR 条目 %d 个: %s' % (len(removed), removed), flush=True)

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(OUT)
time.sleep(4)

ms = doc.Tree.FindNode(r'\Data\Properties\Molecular Structure')
names = [ms.Elements.Item(i).Name for i in range(ms.Elements.Count)] if ms else []
print('组分 %d: %s' % (len(names), names), flush=True)

print()
print('=== 试写 S-101 温压 ===')
for p, v, tag in [(r'\Data\Streams\S-101\Input\TEMP', 25.0, 'TEMP'),
                  (r'\Data\Streams\S-101\Input\PRES', 2.0, 'PRES')]:
    try:
        doc.Tree.FindNode(p).Value = v
        print('   %-6s 成功 -> %r' % (tag, doc.Tree.FindNode(p).Value), flush=True)
    except Exception as e:
        print('   %-6s 失败 %s' % (tag, str(e)[:90]), flush=True)

print()
print('=== 估算相关节点 ===')
for path in [r'\Data\Properties\Estimation\Estimate\Input',
             r'\Data\Properties\Parameters\Binary Interaction',
             r'\Data\Properties\Parameters\Binary Interaction\UNIFAC']:
    n = doc.Tree.FindNode(path)
    print('  %-52s %s' % (path, 'None' if n is None else '子=%s' % n.Elements.Count), flush=True)

print()
print('=== 运行 ===')
try:
    doc.Engine.Run2()
    print('  Run2 返回正常', flush=True)
except Exception as e:
    print('  Run2 抛错:', str(e)[:110], flush=True)

print()
print('=== 找控制面板/消息节点 ===')
d = doc.Tree.FindNode(r'\Data')
for i in range(d.Elements.Count):
    print('   ', d.Elements.Item(i).Name, flush=True)
try:
    doc.SaveAs(r'D:\<化工工作区>\_probe\noair_run.bkp')
except Exception:
    pass
try:
    doc.Close()
except Exception:
    pass
