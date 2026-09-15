# -*- coding: utf-8 -*-
"""BIP 攻关 v2：按出厂示例格式补 FILE-SYM-NAM，测 4库/6库 两档"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>'
SRC = os.path.join(BASE, 'NA_work.bkp')
t0 = open(SRC, encoding='utf-8', errors='ignore').read()

OLD = r'\ DATABANKS \ ? COMPONENTS MAIN ?'
i = t0.find(OLD)
print('锚点位置:', i, '| 源大小:', len(t0))

DB4 = ('FILE-SYM-NAM = ( "APV150 PURE41" "APV150 AQUEOUS" "APV150 SOLIDS" \n'
       '"APV150 INORGANIC" ) ')
DB6 = ('FILE-SYM-NAM = ( "APV150 PURE41" "APV150 AQUEOUS" "APV150 SOLIDS" \n'
       '"APV150 INORGANIC" "APESV150 AP-EOS" "NISTV150 NIST-TRC" ) ')

variants = {}
if i > 0:
    for tag, db in [('4lib', DB4), ('6lib', DB6)]:
        new = r'\ DATABANKS ' + db + r'\ ? COMPONENTS MAIN ?'
        x = t0.replace(OLD, new, 1)
        p = os.path.join(BASE, 'NA_fix_%s.bkp' % tag)
        open(p, 'w', encoding='utf-8', errors='ignore').write(x)
        variants[tag] = p
        print('生成 %s -> %s' % (tag, os.path.basename(p)))
else:
    print('!! 锚点未找到')
    sys.exit(1)


def count_in_file(p):
    x = open(p, encoding='utf-8', errors='ignore').read()
    j = x.find('PARAMNAME = NRTL')
    if j < 0:
        return 0, [], False
    seg = x[j:j + 60000]
    k = seg.find('\n\\ ? ')
    if k > 0:
        seg = seg[:k]
    c = re.findall(r'CID1\s*=\s*"?([A-Za-z0-9_-]+)"?\s+CID2\s*=\s*"?([A-Za-z0-9_-]+)"?', seg)
    sym = 'FILE-SYM-NAM' in x
    return len(c), ['%s-%s' % t for t in c], sym


def run(path, tag, estimate):
    print()
    print('=' * 64)
    print('变体 %s | ESTIMATE=%s' % (tag, estimate))
    msgs = []

    class Sink:
        def OnControlPanelMessage(self, *a):
            s = ' '.join(str(x) for x in a).strip()
            if s and s != 'False':
                msgs.append(s)

    # 若需估算，先改 ESTIMATE
    src = path
    if estimate:
        x = open(path, encoding='utf-8', errors='ignore').read()
        x, n = re.subn(r'ESTIMATE = (?:NO|YES)', 'ESTIMATE = YES', x)
        src = os.path.join(BASE, 'NA_fix_%s_est.bkp' % tag)
        open(src, 'w', encoding='utf-8', errors='ignore').write(x)
        print('  ESTIMATE->YES:', n)

    doc = win32.DispatchEx('Apwn.Document')
    doc.SuppressDialogs = True
    win32.WithEvents(doc, Sink)
    doc.InitFromArchive2(src)
    time.sleep(3)
    doc.Engine.Run2(False)
    t = time.time()
    while time.time() - t < 420:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.25)
        if time.time() - t > 10:
            try:
                if not bool(doc.Engine.IsRunning):
                    break
            except Exception:
                break
    for _ in range(80):
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)

    est_fail = [m for m in msgs if 'CANNOT BE ESTIMATED' in m]
    print('  面板消息:', len(msgs), '| UNIFAC 估算失败:', len(est_fail))
    for m in msgs[-8:]:
        if re.search(r'ERROR|Summary|Terminal|Severe|Errors|CANNOT', m):
            print('   >', m[:110])

    out = os.path.join(BASE, 'NA_fixres_%s%s.bkp' % (tag, '_est' if estimate else ''))
    try:
        doc.SaveAs(out)
    except Exception as e:
        print('  SaveAs 失败:', str(e)[:70])
    try:
        doc.Close()
    except Exception:
        pass
    return out


res = {}
for tag, p in variants.items():
    res[tag] = run(p, tag, estimate=False)
res['4lib_est'] = run(variants['4lib'], '4lib', estimate=True)

print()
print('=' * 64)
print('最终统计（输入文件 vs Aspen 保存后）')
for tag, p in variants.items():
    n, pairs, sym = count_in_file(p)
    print('  输入 %-8s sym=%-5s pairs=%d' % (tag, sym, n))
for tag, p in res.items():
    if p and os.path.exists(p):
        n, pairs, sym = count_in_file(p)
        print('  输出 %-12s sym=%-5s pairs=%d  %s' % (tag, sym, n, ', '.join(pairs[:6])))
print('DONE')
