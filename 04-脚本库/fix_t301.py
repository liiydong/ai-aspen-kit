# -*- coding: utf-8 -*-
"""把 T-301 的第三个进料挪走（Extract 只支持 2 个进料口），再运行"""
import time, re
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\subs1.bkp'
OUT = r'D:\<化工工作区>\_probe\fix301.bkp'

t = open(SRC, encoding='utf-8', errors='ignore').read()


def repl_in(text, blkid, new_in):
    """替换某 BLOCK 的 IN 列表"""
    m = re.search(r'BLOCK\s+BLKID\s*=\s*"%s"' % re.escape(blkid), text)
    if not m:
        print('  未找到 %s' % blkid, flush=True)
        return text
    m2 = re.search(r'IN\s*=\s*\(', text[m.end():])
    if not m2:
        print('  %s 无 IN(' % blkid, flush=True)
        return text
    s = m.end() + m2.end()
    e = text.find(')', s)
    old = text[s:e]
    print('  %s 原 IN:%s' % (blkid, re.sub(r'\s+', ' ', old)), flush=True)
    print('  %s 新 IN:%s' % (blkid, re.sub(r'\s+', ' ', new_in)), flush=True)
    return text[:s] + new_in + text[e:]


t = repl_in(t, 'T-301', ' "S-112" M0-1 "S-115" M1-2 ')
t = repl_in(t, 'T-401', ' "S-114" M0-1 "S-111" M0-1 ')
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)
print('已写出', OUT, flush=True)

# 顺带检查所有 BLOCK 的 IN 个数
print()
print('=== 各 BLOCK 的进料数 ===', flush=True)
for m in re.finditer(r'BLOCK\s+BLKID\s*=\s*"([^"]+)"[^;]*?IN\s*=\s*\(([^)]*)\)', t):
    bid = m.group(1)
    n = len(re.findall(r'M\d+-\d+', m.group(2)))
    lim = {'T-301': 2}.get(bid, '多')
    flag = '⚠' if (bid == 'T-301' and n > 2) else ''
    print('  %-7s %d 个进料 %s' % (bid, n, flag), flush=True)

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(OUT)
time.sleep(4)
print()
print('打开后 \\Data 子节点:', doc.Tree.FindNode(r'\Data').Elements.Count, flush=True)
bl = doc.Tree.FindNode(r'\Data\Blocks')
print('模块数:', bl.Elements.Count, flush=True)
st = doc.Tree.FindNode(r'\Data\Streams')
print('流股数:', st.Elements.Count, flush=True)
t0 = time.time()
try:
    doc.Engine.Run2()
except Exception as e:
    print('Run2 异常', str(e)[:120], flush=True)
while time.time() - t0 < 150:
    time.sleep(3)
    try:
        if doc.Engine.IsRunning is False:
            break
    except Exception:
        break
print('运行用时 %.0fs' % (time.time() - t0), flush=True)
print()
got = []
for b in ['M-101', 'E-101', 'R-101', 'T-201', 'T-301', 'T-401', 'T-404', 'R-501', 'E-601']:
    n = doc.Tree.FindNode(r'\Data\Blocks\%s\Output' % b)
    if n is None:
        got.append('%s:无' % b)
        continue
    vals = []
    for f in ['TEMP', 'PRES', 'DUTY', 'MOLE_RR']:
        try:
            x = n.Elements.Item(f)
            if x is not None and x.Value not in (None, ''):
                vals.append('%s=%s' % (f, x.Value))
        except Exception:
            pass
    got.append('%s[%s]' % (b, ','.join(vals) if vals else '空'))
print('  ', ' '.join(got), flush=True)
try:
    doc.SaveAs(r'D:\<化工工作区>\_probe\fix301_saved.bkp')
except Exception:
    pass
try:
    doc.Close()
except Exception:
    pass
