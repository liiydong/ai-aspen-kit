# -*- coding: utf-8 -*-
"""彻底删除 AIR 组分（COMPONENTS 条目 + MOLEC-STRUCT 段），再运行"""
import time, re
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\full3.bkp'
OUT = r'D:\<化工工作区>\_probe\noair2.bkp'

t = open(SRC, encoding='utf-8', errors='ignore').read()

# 1) 删 COMPONENTS 段里的 AIR 条目
start = t.find('? COMPONENTS MAIN ?')
end = t.find('\\ ? COMPONENTS "ADA/PCS"', start)
seg = t[start:end]
parts = re.split(r'(\s/\s)', seg)
kept = []
n_rm = 0
for p in parts:
    if re.match(r'^\s/\s$', p):
        kept.append(p)
    elif re.search(r'CID = AIR\b', p):
        n_rm += 1
        # 丢弃该条目（连同它前面的分隔符也要清掉）
        if kept and re.match(r'^\s/\s$', kept[-1]):
            kept.pop()
    else:
        kept.append(p)
newseg = ''.join(kept)
t = t[:start] + newseg + t[end:]
print('删除 COMPONENTS 里的 AIR 条目:', n_rm, flush=True)

# 2) 删 ? PROPERTIES "MOLEC-STRUCT" AIR ? 段
m = re.search(r'\?\s*PROPERTIES\s+"MOLEC-STRUCT"\s+AIR\s*\?', t)
if m:
    nxt = t.find('?', m.end())
    # 找到下一个段标记起点
    m2 = re.search(r'\?\s+(PROPERTIES|POLYMERS|"PROP-SET"|"STREAM|BLOCK|STREAM|REACTIONS)',
                   t[m.end():])
    cut_end = m.end() + m2.start() if m2 else m.end() + 40
    t = t[:m.start()] + t[cut_end:]
    print('删除 MOLEC-STRUCT AIR 段', flush=True)
else:
    print('未找到 MOLEC-STRUCT AIR 段', flush=True)

open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(OUT)
time.sleep(4)
ms = doc.Tree.FindNode(r'\Data\Properties\Molecular Structure')
names = [ms.Elements.Item(i).Name for i in range(ms.Elements.Count)] if ms else []
print('组分数 %d: %s' % (len(names), names), flush=True)
print('打开 Ready =', doc.Engine.Ready, flush=True)

t0 = time.time()
try:
    doc.Engine.Run2()
except Exception as e:
    print('Run2 异常', str(e)[:120], flush=True)
while time.time() - t0 < 120:
    time.sleep(3)
    try:
        if doc.Engine.IsRunning is False:
            break
    except Exception:
        break
print('运行用时 %.0fs' % (time.time() - t0), flush=True)
print()
for b in ['M-101', 'E-101', 'R-101', 'T-201', 'T-404', 'R-501']:
    n = doc.Tree.FindNode(r'\Data\Blocks\%s\Output' % b)
    if n is None:
        print('  %-7s 无 Output' % b, flush=True)
        continue
    info = []
    for f in ['TEMP', 'PRES', 'DUTY', 'MOLE_RR', 'ERR']:
        try:
            x = n.Elements.Item(f)
            if x is not None and x.Value not in (None, ''):
                info.append('%s=%r' % (f, x.Value))
        except Exception:
            pass
    print('  %-7s %s' % (b, '; '.join(info) if info else '(空)'), flush=True)

# 流股输出（试多个路径）
print()
for s in ['S-104', 'S-106', 'S-110', 'S-201']:
    for p in ['\\Data\\Streams\\%s\\Output\\TEMP' % s,
              '\\Data\\Streams\\%s\\Output\\B_TEMP' % s,
              '\\Data\\Streams\\%s\\Output\\MASSFLMX' % s]:
        n = doc.Tree.FindNode(p)
        if n is not None and n.Value not in (None, ''):
            print('  %s' % p, '=', n.Value, flush=True)
try:
    doc.SaveAs(r'D:\<化工工作区>\_probe\noair2_saved.bkp')
except Exception:
    pass
try:
    doc.Close()
except Exception:
    pass
