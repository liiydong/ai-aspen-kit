# -*- coding: utf-8 -*-
"""把从 Aspen 出厂示例中提取的真实 NRTL 二元参数注入模型，并跑通验证"""
import os, re, time, json
import win32com.client as w32

BASE = r'D:\<化工工作区>\NA_struct.bkp'
OUT = r'D:\<化工工作区>\_probe'

# 真实参数（来源：Aspen Plus V15 出厂示例文件，即内置物性库检索结果）
PAIRS = [
    # (CID1, CID2, [uvals...], 来源)
    ('O2', 'H2O', [-3.28368021, 17.8247855, 1617.10561, -381.901258, 0.3,
                   0, 0, 0, 0, 0, 0, 1000], 'Industrial Scale Alkaline Electrolyzer.bkp'),
    ('H2O', 'CO2', [10.0640, 10.0640, -3268.1350, -3268.1350, .20], 'flue_gas.bkp'),
    ('H2O', 'HCN', [.0, .0, 505.50, .0, .30], 'enh3hc.bkp'),
    ('H2O', 'NH3', [-.5440720, -.16424220, 1678.4690, -1027.5250, .20], 'enh3co.bkp'),
    ('H2O', 'TOL', [627.05280, -247.87920, -27269.360, 14759.760, .20,
                    0.0, -92.71820, 35.5820, 0.0, 0.0, 264.150, 366.150], 'glycols.bkp'),
]


def fmt(v):
    s = ('%.8f' % v).rstrip('0').rstrip('.')
    return s if s not in ('', '-0') else '0.0'


def rec(c1, c2, vals):
    parts = []
    for k, v in enumerate(vals):
        parts.append('UVAL%d = %s <0> <0>' % (k + 1, fmt(v)))
    return ('\\ BPVAL PARAMNAME2 = NRTL CID1 = %s CID2 = %s UNITROW2 = 0 '
            'TUNITROW2 = 22 TUNITLABEL2 = K %s \\' % (c1, c2, ' '.join(parts)))


def wrap(s, width=76):
    """按 Aspen 惯例在 ~76 字符处折行（折行处插入换行，不在单词中间截断）"""
    out, line = [], ''
    for tok in s.split(' '):
        if len(line) + len(tok) + 1 > width:
            out.append(line)
            line = tok
        else:
            line = (line + ' ' + tok) if line else tok
    if line:
        out.append(line)
    return '\n'.join(out)


t = open(BASE, encoding='utf-8', errors='ignore').read()

# 去掉原空壳记录（VAL 只有来源标记、无 UVAL）
shell = re.search(r'\\ BPVAL PARAMNAME2 = NRTL CID1 =\s*H2O CID2 = TOL.*?VAL12 = "[^"]*"\s*\\', t, re.S)
print('找到空壳记录:', bool(shell))

# 用真实记录替换
newrecs = '\n'.join(wrap(rec(*p[:3])) for p in PAIRS)
if shell:
    t = t[:shell.start()] + newrecs + t[shell.end():]
else:
    # 插到 NRTL-1 段头之后（ESTIMATE = NO 之后）
    m = re.search(r'(ESTIMATE = NO\s*\\)', t)
    print('备用插点:', bool(m))
    t = t[:m.end()] + '\n' + newrecs + t[m.end():]

p = os.path.join(OUT, 'inj.bkp')
open(p, 'w', encoding='utf-8', errors='ignore').write(t)
print('已写出', p, len(t))

# 跑通并验证
sp = os.path.join(OUT, 'inj_s.bkp')
app = w32.DispatchEx('Apwn.Document')
try:
    app.InitFromArchive2(p)
    print('\\Data children =', app.Tree.FindNode(r'\Data').Elements.Count)
    app.Engine.Run2(False)
    time.sleep(4)
    app.SaveAs(sp)
    try:
        app.Close()
    except Exception:
        pass
    try:
        app.Quit()
    except Exception:
        pass
except Exception as e:
    print('RUN ERR', str(e)[:150])

s = open(sp, encoding='utf-8', errors='ignore').read() if os.path.exists(sp) else ''
i = s.find('PARAMNAME = NRTL')
j = s.find('PARAMNAME =', i + 20)
seg = s[i:j] if j > i else s[i:i+40000]
print('保存文件: UVAL=%d BPVAL=%d size=%d' % (seg.count('UVAL'), seg.count('BPVAL'), len(s)))
cids = re.findall(r'CID1 = ("?[^"\s]+"?)\s+CID2 = ("?[^"\s]+"?)', seg)
print('参数对:', cids)
