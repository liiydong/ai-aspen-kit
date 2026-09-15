# -*- coding: utf-8 -*-
"""Q3：T-403 改 Sep（只改 T-403 自己的 FLOWSHEET 条目，不误伤 T-404）"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_最终版.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_q3.bkp'
t = open(SRC, encoding='utf-8', errors='ignore').read()
assert len(re.findall(r'\?\s*BLOCK\s+RADFRAC', t)) == 5

# 1) 注册表
assert 'T-403\nRadFrac\nBuilt-In\nRADFRAC\n' in t
t = t.replace('T-403\nRadFrac\nBuilt-In\nRADFRAC\n', 'T-403\nSep\nBuilt-In\nSEP\n', 1)
print('1) 注册表 ✓')

# 2) FLOWSHEET：只改 T-403 自己的条目
i = t.find('BLKID = "T-403"')
j = t.find('BLOCK', i + 20)
seg = t[i:j]
print('   原:', repr(seg))
seg = seg.replace('"RADFRAC"', '"SEP"').replace('"RadFrac"', '"Sep"')
seg = seg.replace('M1-2', 'M0-1').replace('M2-3', 'M0-1')
t = t[:i] + seg + t[j:]
print('   新:', repr(seg))
# 校验 T-404 未被误改
m404 = re.search(r'BLKID = "T-404"[^\\]{0,160}', t)
print('   T-404 现状:', repr(m404.group()[:150]) if m404 else '未找到')
assert 'M1-2' in m404.group(), 'T-404 端口被误改'
print('2) FLOWSHEET ✓（T-404 未受影响）')

# 3) 块段落
def wr(s, w=74):
    out, line = [], ''
    for tok in s.split(' '):
        if len(line) + len(tok) + 1 > w:
            out.append(line)
            line = tok
        else:
            line = (line + ' ' + tok) if line else tok
    if line:
        out.append(line)
    return '\n'.join(out)


ci = t.find('? COMPONENTS MAIN ?')
cj = t.find('? COMPONENTS "COMP-LIST"', ci)
cids = re.findall(r'CID = "?([A-Za-z0-9-]+)"?', t[ci:cj])
FR = {'4-CP': 0.99, '3-CP': 0.0002}
recs = ['PARAM-STREAM = "S-119" SUBSTREAM = MIXED COMPS = "%s" FRACS = %s <0> <0> '
        % (c, FR.get(c, 0.99)) for c in cids]
para = ('? BLOCK SEP "T-403" ? ; "METCBAR_MOLE" ; ; ICON1 ; \\ '
        'DESCRIPTION DESCRIPTION = "4-cyanopyridine removal, 4-CP overhead per Ruibang" \\ \\ '
        'PARAM1 PRES1 = 0.40 <20> <5> \\ \\ PARAM ' + ' /  '.join(recs) + ' \\ ')
starts = [s.start() for s in re.finditer(r'\?\s*BLOCK\s+', t)]
done = False
for k in range(len(starts)):
    s0 = starts[k]
    s1 = starts[k + 1] if k + 1 < len(starts) else len(t)
    m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
    if m and m.group(2) == 'T-403':
        t = t[:s0] + wr(para) + t[s1:]
        done = True
        break
assert done
print('3) T-403 段 -> SEP ✓')

assert len(re.findall(r'\?\s*BLOCK\s+RADFRAC', t)) == 4
assert len(re.findall(r'BLOCK\s+RADFRAC\s*\n?\s*"T-401"', t)) == 1
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)
print('写出:', OUT, os.path.getsize(OUT))

MSGS = []


class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            MSGS.append(s)


doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
try:
    win32.WithEvents(doc, Sink)
except Exception:
    pass
doc.InitFromArchive2(OUT)
time.sleep(3)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 320:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 10:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(50):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)


def g(p):
    n = doc.Tree.FindNode(p)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return None


def comp(s, cid):
    n = doc.Tree.FindNode(r'\Data\Streams\%s\Output\MASSFLOW3' % s)
    if n is None:
        return None
    try:
        for i in range(n.Elements.Count):
            e = n.Elements.Item(i)
            if e.Name.upper() == cid.upper():
                return float(e.Value) if e.Value else 0.0
    except Exception:
        pass
    return None


print()
print('=== 错误汇总 ===')
for x in MSGS:
    if 'Summary' in x or 'Errors' in x or 'Terminal' in x or 'Severe' in x or 'completed' in x:
        print('   |', x[:160])
print()
for s in ['S-118', 'S-119', 'S-120', 'S-121', 'S-122']:
    print('  %s 总=%-10s 3-CP=%-9s 4-CP=%-7s' % (
        s, g(r'\Data\Streams\%s\Output\MASSFLMX\MIXED' % s), comp(s, '3-CP'), comp(s, '4-CP')))
try:
    doc.SaveAs(OUT.replace('.bkp', '_s.bkp'))
    print('已保存')
except Exception as ex:
    print('保存失败:', ex)
try:
    doc.Close()
except Exception:
    pass
print('DONE')
