# -*- coding: utf-8 -*-
"""解析 Aspen xlsx 报告（分列 tsv）→ 结构化流股/模块数据 JSON。"""
import re, sys, io, json, collections
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

SRC = r'D:\<化工工作区>\_probe\report_cols.tsv'
OUT = r'D:\<化工工作区>\_probe\report_parsed.json'

COMPONENTS = ['3-MP','NH3','O2','N2','3-CP','4-CP','H2O','TOL','CO2','HCN','CO','AIR','NAM','NAC','4-MP']

NUM = re.compile(r'[-+]?(?:\d+\.\d+|\d+)(?:[-+]\d+)?')
STREAMID = re.compile(r'S-\d+[A-Z]?')

def nums(s, n):
    """取行中最后 n 个数值 token"""
    toks = NUM.findall(s)
    if len(toks) < n:
        return None
    return [float_asp(x) for x in toks[-n:]]

def float_asp(t):
    """Aspen 数值：1.2345-02 → 0.012345"""
    m = re.match(r'^([-+]?)(\d+\.\d+|\d+)(?:([-+])(\d+))?$', t)
    if not m:
        return None
    sign, mant, esign, exp = m.groups()
    v = float(mant)
    if esign:
        v = v * (10 ** (-int(exp) if esign == '-' else int(exp)))
    return -v if sign == '-' else v

# ---- 读入并拼接 ----
# Aspen 导出的 xlsx 把每行按固定字符位切成 A/B/C 三列，切点处会丢空格：
#   · B→C 若 B 尾不是 +/-，需补一个空格（否则两个数值粘连，如 13.905913.9059）
#   · B 尾为 -/+ 时说明该符号属于 C 中数值（如 -0.1482 + -6.0187-02），不补
raw = []
for ln in open(SRC, encoding='utf-8').read().split('\n'):
    if not ln.strip():
        continue
    p = ln.split('\t')
    r = int(p[0]); cols = (p[1:] + ['', '', ''])[:3]
    a, b, c = cols[0], cols[1], cols[2]
    if b and c and not b.endswith(('-', '+')):
        b = b + ' '
    raw.append((r, (a + b + c).rstrip()))

streams = collections.OrderedDict()
blocks = collections.OrderedDict()

mode = None           # 'stream' | 'block'
cur_ids = []          # 当前页流股
cur_comp = None
page = None

PAGEHDR = re.compile(r'^S-\d+[A-Z]?(?:\s+S-\d+[A-Z]?)*$')

for r, s in raw:
    st = s.strip()
    if PAGEHDR.match(st) and 'S-' in st and len(st) < 60:
        ids = STREAMID.findall(st)
        if len(ids) >= 1:
            cur_ids = ids
            for sid in ids:
                streams.setdefault(sid, {})
            continue
    if st.startswith('BLOCK:'):
        mode = 'block'; continue

    n = len(cur_ids)
    if n == 0:
        continue

    # 组分行
    for comp in COMPONENTS:
        if st.startswith(comp) and (len(st) == len(comp) or not st[len(comp)].isalnum()):
            seg = st[len(comp):]
            v = nums(seg, n) or nums(s, n)
            if v:
                for i, sid in enumerate(cur_ids):
                    streams[sid].setdefault('comp_kmol', {})[comp] = v[i]
            break
    else:
        # 状态变量等
        key = None
        for tag, name in [
            ('TEMP', 'TEMP'), ('PRES', 'PRES'), ('VFRAC', 'VFRAC'),
            ('LFRAC', 'LFRAC'), ('SFRAC', 'SFRAC'),
            ('KCAL/MOL', 'H_KCALMOL'), ('KCAL/KG', 'H_KCALKG'), ('GCAL/HR', 'H_GCALHR'),
            ('CAL/MOL-K', 'S_CALMOLK'), ('CAL/GM-K', 'S_CALGMK'),
            ('KMOL/CUM', 'RHO_KMOLCUM'), ('KG/CUM', 'RHO_KGCUM'),
            ('AVG MW', 'MW'),
            ('KMOL/HR', 'F_KMOL'), ('KG/HR', 'F_KG'), ('CUM/HR', 'F_CUM'),
        ]:
            if st.startswith(tag) and (len(st) == len(tag) or not st[len(tag)].isalnum()):
                if tag in ('KCAL/KG', 'KG/CUM', 'KCAL/MOL', 'KMOL/CUM', 'KMOL/HR', 'KG/HR'):
                    pass
                key = name
                break
        if key:
            v = nums(s, n)
            if v:
                for i, sid in enumerate(cur_ids):
                    streams[sid][key] = v[i]
        # 相态
        if st.startswith('PHASE:'):
            tail = st[6:]
            ph = tail.split()
            if len(ph) == n:
                for i, sid in enumerate(cur_ids):
                    streams[sid]['PHASE'] = ph[i]

json.dump({'streams': streams, 'blocks': blocks}, open(OUT, 'w', encoding='utf-8'),
          ensure_ascii=False, indent=1)

msg = ['parsed streams: %d' % len(streams)]
for sid, d in streams.items():
    msg.append('%s comp=%d keys=%s' % (sid, len(d.get('comp_kmol', {})), ','.join(sorted(d.keys()))))
print('\n'.join(msg))
