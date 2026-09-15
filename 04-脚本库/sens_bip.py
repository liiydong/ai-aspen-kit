# -*- coding: utf-8 -*-
"""敏感性分析：删掉唯一 1 对 NRTL 参数（全 0）vs 现状，看关键结果变化"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>'
SRC = os.path.join(BASE, 'NA_work.bkp')
t0 = open(SRC, encoding='utf-8', errors='ignore').read()

# 定位并删除 H2O-TOL 的 BPVAL 记录
i = t0.find('BPVAL PARAMNAME2 = NRTL CID1 = H2O CID2 = TOL')
print('H2O-TOL 记录位置:', i)
va = t0
if i > 0:
    # 往前找到记录起始的 '\\ \\ '
    j = t0.rfind('\\ \\ BPVAL', 0, i + 10)
    if j < 0:
        j = t0.rfind('\\', 0, i) - 1
    # 往后找到 ' / ' 结束
    k = t0.find(' / ', i)
    print('  起点 %d 终点 %d' % (j, k))
    if j > 0 and k > j:
        va = t0[:j] + t0[k + 3:]
        print('  已删除 %d 字符' % (k + 3 - j))
V = os.path.join(BASE, 'NA_nobip.bkp')
open(V, 'w', encoding='utf-8', errors='ignore').write(va)

KEYS = ['S-209', 'S-114', 'S-113', 'S-121', 'S-117', 'S-122', 'S-119', 'S-120']


def run(path, tag):
    msgs = []

    class Sink:
        def OnControlPanelMessage(self, *a):
            s = ' '.join(str(x) for x in a).strip()
            if s and s != 'False':
                msgs.append(s)

    doc = win32.DispatchEx('Apwn.Document')
    doc.SuppressDialogs = True
    win32.WithEvents(doc, Sink)
    doc.InitFromArchive2(path)
    time.sleep(3)
    doc.Engine.Run2(False)
    t = time.time()
    while time.time() - t < 420:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.3)
        if time.time() - t > 10:
            try:
                if not bool(doc.Engine.IsRunning):
                    break
            except Exception:
                break
    for _ in range(100):
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)

    errs = [m for m in msgs if re.search(r'Terminal Errors|Severe Errors|  Errors', m)]
    print('===== %s =====' % tag)
    for m in errs[-4:]:
        print('   ', m[:90])

    def g(p):
        n = doc.Tree.FindNode(p)
        if n is None:
            return None
        try:
            return n.Value
        except Exception:
            return None

    print('  %-8s %-14s %-12s %-12s' % ('流股', 'MASS kg/h', 'TEMP C', 'PRES bar'))
    for s in KEYS:
        b = r'\Data\Streams\%s\Output' % s
        m = g(b + r'\MASSFLMX\MIXED')
        tp = g(b + r'\TEMP_OUT\MIXED')
        pp = g(b + r'\PRES_OUT\MIXED')
        try:
            m = '%.2f' % float(m)
        except Exception:
            m = str(m)
        try:
            tp = '%.2f' % float(tp)
        except Exception:
            tp = str(tp)
        try:
            pp = '%.4f' % float(pp)
        except Exception:
            pp = str(pp)
        print('  %-8s %-14s %-12s %-12s' % (s, m, tp, pp))

    # 塔负荷
    print('  --- 塔 ---')
    for b in ['T-401', 'T-402', 'T-403', 'T-404', 'T-301', 'T-201']:
        q = g(r'\Data\Blocks\%s\Output\COND_DUTY' % b)
        q2 = g(r'\Data\Blocks\%s\Output\REB_DUTY' % b)
        try:
            q = '%.1f' % float(q)
        except Exception:
            q = str(q)
        try:
            q2 = '%.1f' % float(q2)
        except Exception:
            q2 = str(q2)
        print('    %-7s COND=%-12s REB=%-12s' % (b, q, q2))

    # 纯度
    nd = doc.Tree.FindNode(r'\Data\Streams\S-209\Output\MASSFLOW3')
    if nd is not None:
        try:
            for i in range(nd.Elements.Count):
                e = nd.Elements.Item(i)
                try:
                    v = float(e.Value or 0)
                    if v > 0.5:
                        print('    S-209 组分 %-8s %8.2f' % (e.Name, v))
                except Exception:
                    pass
        except Exception:
            pass

    try:
        doc.Close()
    except Exception:
        pass


run(SRC, '基线（1 对参数 H2O-TOL）')
run(V, '变体（0 对参数，全理想溶液）')
print('DONE')
