# -*- coding: utf-8 -*-
"""探测：FLOWSHEET 原文、流股 Port、块 Output 子节点、S-207"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_最终定稿.bkp'
t = open(SRC, encoding='utf-8', errors='ignore').read()
L = []

# 1) FLOWSHEET 段原文
L.append('===== FLOWSHEET 出现位置 =====')
for m in re.finditer(r'FLOWSHEET', t):
    i = m.start()
    L.append('  pos %d : %r' % (i, t[i - 20:i + 260]))
    if len(L) > 12:
        break

# 2) 所有段头（前 40 个）
L.append('')
L.append('===== 段头序列（前 45） =====')
for i, m in enumerate(re.finditer(r'\?\s*([A-Z][A-Z0-9 /_.-]{0,44}?)\s*\?', t)):
    L.append('  %7d  %s' % (m.start(), m.group().strip()))
    if i > 44:
        break

# 3) S-207 段原文
L.append('')
L.append('===== S-207 段原文 =====')
for m in re.finditer(r'STREAM\s+MATERIAL\s*\n?\s*"S-207"\s*\?', t):
    i = m.start()
    L.append(re.sub(r'\s+', ' ', t[i:i + 600]))

doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
doc.InitFromArchive2(SRC)
time.sleep(2)

def kids(path, lim=40):
    n = doc.Tree.FindNode(path)
    if n is None:
        return ['<None>']
    out = []
    try:
        for i in range(min(n.Elements.Count, lim)):
            e = n.Elements.Item(i)
            out.append('%s=%s' % (e.Name, str(e.Value)[:40]))
    except Exception as ex:
        out.append('ERR ' + str(ex))
    return out

L.append('')
L.append('===== \\Data\\Streams\\S-101 子节点 =====')
L.append('  ' + ' | '.join(kids(r'\Data\Streams\S-101')))
L.append('===== \\Data\\Streams\\S-101\\Ports =====')
L.append('  ' + ' | '.join(kids(r'\Data\Streams\S-101\Ports')))
L.append('===== \\Data\\Streams\\S-101\\Output 子节点 =====')
L.append('  ' + ' | '.join(kids(r'\Data\Streams\S-101\Output')))
L.append('===== \\Data\\Blocks\\T-401 子节点 =====')
L.append('  ' + ' | '.join(kids(r'\Data\Blocks\T-401')))
L.append('===== \\Data\\Blocks\\T-401\\Output 子节点 =====')
L.append('  ' + ' | '.join(kids(r'\Data\Blocks\T-401\Output')))
L.append('===== \\Data\\Blocks\\T-401\\Input 子节点 =====')
L.append('  ' + ' | '.join(kids(r'\Data\Blocks\T-401\Input')))

open(r'D:\<化工工作区>\_probe\probe_ports.txt', 'w', encoding='utf-8').write('\n'.join(L))
try:
    doc.Close()
except Exception:
    pass
print('DONE')
