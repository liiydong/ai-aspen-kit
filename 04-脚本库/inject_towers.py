# -*- coding: utf-8 -*-
"""注入 5 座 RadFrac 的完整规格，然后运行"""
import time, re
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\tp_ok.bkp'
OUT = r'D:\<化工工作区>\_probe\full1.bkp'

def radfrac(bid, n, feed_sid, feed_stage, dist, bot, p1, df, rr,
            cond='TOTAL', reb='KETTLE', icon='FRACT1'):
    """按参考模型格式拼 RadFrac 段"""
    return ('? BLOCK RADFRAC "%s" ? ; "METCBAR_MOLE" ; ; %s ; \\ '
            'PARAM NSTAGE = %d OPT-PRES = "DP-COL" OPT-PRES-TOP = "DP-COND" NSTAGEMAX = %d '
            '\\ \\ "COL-CONFIG" CONDENSER = %s REBOILER = %s '
            '\\ \\ FEEDS FEED-SID = "%s" FEED-STAGE = %d '
            '\\ \\ PRODUCTS PROD-STREAM = "%s" PROD-STAGE = 1 PROD-PHASE = L P-S = N / '
            'PROD-STREAM = "%s" PROD-STAGE = %d PROD-PHASE = L P-S = N '
            '\\ \\ "P-SPEC2" PRES1 = %s <20> <10> '
            '\\ \\ "COL-SPECS" D:F = %s <-1> <0> BASIS-RDV = 0.0 <-1> <0> BASIS-RR = %s <-1> <0> '
            '\\ \\ "KLL-VECS" \\ \\ "TRSZ-VECS" \\ \\ "PCKSR-VECS" \\ '
            % (bid, icon, n, n + 1, cond, reb, feed_sid, feed_stage,
               dist, bot, n, p1, df, rr))

# T-201 吸收塔：无冷凝器/再沸器，两个进料
T201 = ('? BLOCK RADFRAC "T-201" ? ; "METCBAR_MOLE" ; ; FRACT1 ; \\ '
        'PARAM NSTAGE = 6 OPT-PRES = "DP-COL" OPT-PRES-TOP = "DP-COND" NSTAGEMAX = 7 '
        '\\ \\ "COL-CONFIG" CONDENSER = NONE REBOILER = NONE '
        '\\ \\ FEEDS FEED-SID = "S-107" FEED-STAGE = 1 / FEED-SID = "S-108" FEED-STAGE = 6 '
        '\\ \\ PRODUCTS PROD-STREAM = "S-109" PROD-STAGE = 1 PROD-PHASE = V P-S = N / '
        'PROD-STREAM = "S-110" PROD-STAGE = 6 PROD-PHASE = L P-S = N '
        '\\ \\ "P-SPEC2" PRES1 = 1.0 <20> <10> '
        '\\ \\ "KLL-VECS" \\ \\ "TRSZ-VECS" \\ \\ "PCKSR-VECS" \\ ')

NEW = {
    'T-201': T201,
    'T-401': radfrac('T-401', 25, 'S-114', 12, 'S-115', 'S-116', 1.0, 24.5, 2.0),
    'T-402': radfrac('T-402', 40, 'S-116', 20, 'S-117', 'S-118', 1.0, 5.0, 3.0),
    'T-403': radfrac('T-403', 45, 'S-118', 22, 'S-119', 'S-120', 0.6, 0.5, 18.0),
    'T-404': radfrac('T-404', 32, 'S-120', 18, 'S-121', 'S-122', 0.3, 9.5, 3.0),
}

t = open(SRC, encoding='utf-8', errors='ignore').read()

# 定位所有 BLOCK 段的起点
starts = [(m.start(), m.group()) for m in
          re.finditer(r'\?\s*BLOCK\s+[A-Z0-9]+\s+"?[A-Za-z0-9-]+"?\s*\?', t)]
print('找到 %d 个 BLOCK 段:' % len(starts), flush=True)

# 反序替换，避免位移
replaced = []
for idx in range(len(starts) - 1, -1, -1):
    s0 = starts[idx][0]
    s1 = starts[idx + 1][0] if idx + 1 < len(starts) else len(t)
    seg = t[s0:s1]
    m = re.match(r'\?\s*BLOCK\s+RADFRAC\s+"([^"]+)"\s*\?', seg)
    if not m:
        continue
    bid = m.group(1)
    if bid in NEW:
        t = t[:s0] + NEW[bid] + '\n' + t[s1:]
        replaced.append(bid)
print('已替换:', replaced, flush=True)
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t)

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive(OUT)
time.sleep(4)

print()
print('=== 读回塔参数 ===', flush=True)
for b in ['T-201', 'T-401', 'T-402', 'T-403', 'T-404']:
    row = []
    for f in ['NSTAGE', 'PRES1']:
        n = doc.Tree.FindNode(r'\Data\Blocks\%s\Input\%s' % (b, f))
        row.append('%s=%s' % (f, n.Value if n is not None else 'None'))
    print('  %-7s %s' % (b, '  '.join(row)), flush=True)

# 把 T-301 压力也试一下
for f in ['PRES', 'PRES1', 'PSTAGE']:
    try:
        doc.Tree.FindNode(r'\Data\Blocks\T-301\Input\%s' % f).Value = 1.0
        print('  T-301.%s 写入成功' % f, flush=True)
        break
    except Exception:
        pass

print()
print('=== 运行 ===', flush=True)
try:
    doc.Engine.Run2()
    print('  Run2 完成', flush=True)
except Exception as e:
    print('  Run2 抛出:', str(e)[:150], flush=True)

print()
print('=== 看结果 ===', flush=True)
for b in ['R-101', 'R-501', 'T-404']:
    for f in ['TEMP', 'PRES']:
        try:
            n = doc.Tree.FindNode(r'\Data\Blocks\%s\Output\%s' % (b, f))
            print('  %s.Output.%s = %r' % (b, f, n.Value if n is not None else None), flush=True)
        except Exception:
            pass
for s in ['S-106', 'S-201', 'S-301', 'S-121']:
    try:
        n = doc.Tree.FindNode(r'\Data\Streams\%s\Output\TEMP' % s)
        print('  %s.Output.TEMP = %r' % (s, n.Value if n is not None else None), flush=True)
    except Exception:
        pass
try:
    doc.SaveAs(r'D:\<化工工作区>\_probe\full1_saved.bkp')
    print('已保存 full1_saved.bkp', flush=True)
except Exception as e:
    print('SaveAs 失败', str(e)[:80], flush=True)
try:
    doc.Close()
except Exception:
    pass
