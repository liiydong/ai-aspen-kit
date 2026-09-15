# -*- coding: utf-8 -*-
"""对照：空壳记录 vs 真实文献参数，看结果是否变化"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>'
SRC = os.path.join(BASE, 'NA_work.bkp')
x = open(SRC, encoding='utf-8', errors='ignore').read()

# ---- 完整 dump NRTL 记录 ----
i = x.find('BPVAL PARAMNAME2 = NRTL')
print('记录起点:', i)
seg = x[i:i + 1200]
# 找记录结束：下一个 '\\ ? ' 或 '\n\\ '
m = re.search(r'\\\s*\?\s*[A-Z]', seg)
end_rel = m.start() if m else len(seg)
print('记录结束(相对):', end_rel)
print('--- 记录全文 ---')
print(repr(seg[:end_rel + 20]))
print()

# ---- 构造变体：把空壳 VAL 换成真实数值 ----
# 真实 H2O-TOL NRTL: aij=-5.05265, aji=5.01128, bij=2182.52, bji=473.976, alpha=0.100364
# 对应 UVAL1..UVAL5
rec_end = i + end_rel
rec = x[i:rec_end]
newrec = ('BPVAL PARAMNAME2 = NRTL CID1 = H2O CID2 = TOL UNITROW2 = 0 TUNITROW2 = 22 TUNITLABEL2 = C \n'
          'UVAL1 = -5.05265 <0> <0> UVAL2 = 5.01128 <0> <0> UVAL3 = 2182.52 <0> <0> '
          'UVAL4 = 473.976 <0> <0> UVAL5 = 0.100364 <0> <0> '
          'VAL1 = "0$-5.05265" VAL2 = "0$5.01128" VAL3 = "0$2182.52" VAL4 = "0$473.976" VAL5 = "0$0.100364" '
          'VAL6 = "0$0.0" VAL7 = "0$0.0" VAL8 = "0$0.0" VAL9 = "0$0.0" VAL10 = "0$0.0" ')
x2 = x[:i] + newrec + x[rec_end:]
V = os.path.join(BASE, 'NA_realtol.bkp')
open(V, 'w', encoding='utf-8', errors='ignore').write(x2)
print('变体写出: %d -> %d bytes' % (len(x), len(x2)))


def run(path, tag):
    msgs = []

    class Sink:
        def OnControlPanelMessage(self, *a):
            s = ' '.join(str(y) for y in a).strip()
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

    print()
    print('===== %s =====' % tag)
    for m in msgs:
        if re.search(r'Terminal Errors|Severe Errors|\bErrors\b|Warnings', m):
            print('   ', m[:88])

    def g(p):
        nd = doc.Tree.FindNode(p)
        if nd is None:
            return None
        try:
            return nd.Value
        except Exception:
            return None

    for s in ['S-209', 'S-114', 'S-121']:
        b = r'\Data\Streams\%s\Output' % s
        v = []
        for k in [r'\MASSFLMX\MIXED', r'\TEMP_OUT\MIXED']:
            try:
                v.append('%.3f' % float(g(b + k)))
            except Exception:
                v.append(str(g(b + k)))
        print('  %-7s MASS=%-12s TEMP=%-10s' % (s, v[0], v[1]))
    for b in ['T-401', 'T-402', 'T-404']:
        try:
            print('    %-6s COND=%.3f' % (b, float(g(r'\Data\Blocks\%s\Output\COND_DUTY' % b))))
        except Exception:
            pass
    try:
        doc.Close()
    except Exception:
        pass


run(SRC, '基线（空壳记录）')
run(V, '变体（真实 H2O-TOL 参数）')
print('DONE')
