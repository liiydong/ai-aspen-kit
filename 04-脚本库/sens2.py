# -*- coding: utf-8 -*-
"""严格对照：删除空的 NRTL 记录 vs 保留，比较关键结果"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>'
SRC = os.path.join(BASE, 'NA_work.bkp')
x = open(SRC, encoding='utf-8', errors='ignore').read()

pat = re.compile(r'\\ \\ BPVAL PARAMNAME2 = NRTL CID1 = H2O.*?VAL12 = "APV150 LLE-ASPEN" ', re.S)
x2, n = pat.subn('', x)
print('删除记录数:', n)
V = os.path.join(BASE, 'NA_nobip2.bkp')
open(V, 'w', encoding='utf-8', errors='ignore').write(x2)
print('变体大小: %d -> %d' % (len(x), len(x2)))
i = x2.find('PARAMNAME = NRTL')
print('删除后 NRTL 段尾:', repr(x2[i + 100: i + 340]))


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

    print('  %-7s %-12s %-10s %-10s' % ('流股', 'MASS', 'TEMP', 'PRES'))
    for s in ['S-209', 'S-114', 'S-121', 'S-117', 'S-119']:
        b = r'\Data\Streams\%s\Output' % s
        v = []
        for k in [r'\MASSFLMX\MIXED', r'\TEMP_OUT\MIXED', r'\PRES_OUT\MIXED']:
            try:
                v.append('%.3f' % float(g(b + k)))
            except Exception:
                v.append(str(g(b + k)))
        print('  %-7s %-12s %-10s %-10s' % (s, v[0], v[1], v[2]))

    for b in ['T-401', 'T-402', 'T-404', 'T-201']:
        try:
            q = '%.2f' % float(g(r'\Data\Blocks\%s\Output\COND_DUTY' % b))
        except Exception:
            q = str(g(r'\Data\Blocks\%s\Output\COND_DUTY' % b))
        print('    %-6s COND=%s' % (b, q))
    try:
        doc.Close()
    except Exception:
        pass


run(SRC, '基线（含 1 条空壳记录）')
run(V, '变体（已删除该记录）')
print('DONE')
