# -*- coding: utf-8 -*-
"""测试预测型物性方法 UNIF-LL / UNIFAC 能否让萃取塔分相"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_v13_s.bkp'
base = open(SRC, encoding='utf-8', errors='ignore').read()


def run(method, out):
    t = base
    n = 0
    t, n = re.subn(r'PARAM\s+BASE\s*=\s*[^\s\\]+', 'PARAM BASE = "%s"' % method, t, count=1)
    open(out, 'w', encoding='utf-8', errors='ignore').write(t)
    MSGS = []

    class Sink:
        def OnControlPanelMessage(self, *a):
            s = ' '.join(str(x) for x in a).strip()
            if s and s != 'False':
                MSGS.append(s)

    doc = win32.DispatchEx('Apwn.Document')
    try:
        doc.SuppressDialogs = True
    except Exception:
        pass
    try:
        win32.WithEvents(doc, Sink)
    except Exception:
        pass
    doc.InitFromArchive2(out)
    time.sleep(3)
    doc.Engine.Run2(False)
    t0 = time.time()
    while time.time() - t0 < 300:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.25)
        if time.time() - t0 > 10:
            try:
                if not bool(doc.Engine.IsRunning):
                    break
            except Exception:
                break
    for _ in range(60):
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)

    def g(p):
        n2 = doc.Tree.FindNode(p)
        if n2 is None:
            return None
        try:
            return n2.Value
        except Exception:
            return 'ERR'

    print('=' * 70)
    print('方法 = %s   (替换 %d 次)' % (method, n))
    key = [x for x in MSGS if re.search(r'ERROR|WARNING|BINARY|UNIFAC|UNIF|MISSING|INVALID', x, re.I)]
    print('  关键消息 %d 条:' % len(key))
    for x in key[:12]:
        print('    |', x[:170])
    print('  末尾:')
    for x in MSGS[-6:]:
        print('    |', x[:170])
    print('  --- 萃取相关流股 ---')
    for s in ['S-110', 'S-113', 'S-114']:
        b = r'\Data\Streams\%s\Output' % s
        print('    %s MASS=%s' % (s, g(b + r'\MASSFLMX\MIXED')))
        nn = doc.Tree.FindNode(b + r'\MASSFLOW3')
        if nn is not None:
            try:
                for i2 in range(nn.Elements.Count):
                    e = nn.Elements.Item(i2)
                    try:
                        if e.Value and abs(float(e.Value)) > 0.05:
                            print('        %-8s %9.2f' % (e.Name, float(e.Value)))
                    except Exception:
                        pass
            except Exception:
                pass
    try:
        doc.Close()
    except Exception:
        pass


run('UNIF-LL', r'D:\<化工工作区>\_probe\bip_ll.bkp')
run('UNIFAC', r'D:\<化工工作区>\_probe\bip_fac.bkp')
print('DONE')
