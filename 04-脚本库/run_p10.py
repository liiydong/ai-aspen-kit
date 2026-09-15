# -*- coding: utf-8 -*-
"""P10：T-402 改为脱轻组分（残甲苯 + 未反应 3-MP），恢复产品纯度"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_v20.bkp'
base = open(SRC, encoding='utf-8', errors='ignore').read()


def repl_block(t, bid, fn):
    starts = [s.start() for s in re.finditer(r'\?\s*BLOCK\s+', t)]
    for k in range(len(starts)):
        s0 = starts[k]
        s1 = starts[k + 1] if k + 1 < len(starts) else len(t)
        m = re.match(r'\?\s*BLOCK\s+([A-Z0-9]+)\s+"?([A-Za-z0-9-]+)"?\s*\?', t[s0:s1])
        if m and m.group(2) == bid:
            return t[:s0] + fn(t[s0:s1]) + t[s1:]
    raise RuntimeError(bid)


def build(d402, rr402, d403, rr403, d404, rr404, out):
    t = base
    for bid, df, rr in [('T-402', d402, rr402), ('T-403', d403, rr403), ('T-404', d404, rr404)]:
        def fn(seg, df=df, rr=rr):
            seg = re.sub(r'D:F = [\d.]+ <-1> <0>', 'D:F = %.4f <-1> <0>' % df, seg, count=1)
            if rr:
                seg = re.sub(r'BASIS-RR = [\d.]+ <-1> <0>', 'BASIS-RR = %s <-1> <0>' % rr, seg, count=1)
            return seg
        t = repl_block(t, bid, fn)
    open(out, 'w', encoding='utf-8', errors='ignore').write(t)


def run(tag, path, final=False):
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
    doc.InitFromArchive2(path)
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

    def c(s, cid):
        n = doc.Tree.FindNode(r'\Data\Streams\%s\Output\MASSFLOW3' % s)
        if n is None:
            return 0.0
        try:
            for i in range(n.Elements.Count):
                e = n.Elements.Item(i)
                if e.Name.upper() == cid.upper():
                    return float(e.Value) if e.Value else 0.0
        except Exception:
            pass
        return 0.0

    m121 = g(r'\Data\Streams\S-121\Output\MASSFLMX\MIXED') or 0
    cp = c('S-121', '3-CP')
    print('%-22s S-121=%-8.1f 3-CP=%-8.1f 4-CP=%-6.2f 3-MP=%-6.2f 纯度=%.2f%%' % (
        tag, m121, cp, c('S-121', '4-CP'), c('S-121', '3-MP'),
        (cp / m121 * 100) if m121 > 0 else 0), flush=True)
    if final:
        print()
        print('=== 全塔 ===')
        for b in ['T-201', 'T-401', 'T-402', 'T-403', 'T-404']:
            bb = r'\Data\Blocks\%s\Output' % b
            print('  %-6s COND=%-11s REB=%-11s RR=%-7s Ttop=%-9s Tbot=%-9s' % (
                b, g(bb + r'\COND_DUTY'), g(bb + r'\REB_DUTY'), g(bb + r'\MOLE_RR'),
                g(bb + r'\TOP_TEMP'), g(bb + r'\BOTTOM_TEMP')))
        print()
        print('=== 关键流股 ===')
        for s in ['S-110', 'S-113', 'S-114', 'S-115', 'S-116', 'S-117', 'S-118',
                  'S-119', 'S-120', 'S-121', 'S-122']:
            b = r'\Data\Streams\%s\Output' % s
            print('  --- %s MASS=%s T=%s' % (s, g(b + r'\MASSFLMX\MIXED'), g(b + r'\TEMP_OUT')))
            nn = doc.Tree.FindNode(b + r'\MASSFLOW3')
            if nn is not None:
                for i2 in range(nn.Elements.Count):
                    e = nn.Elements.Item(i2)
                    try:
                        if e.Value and abs(float(e.Value)) > 0.05:
                            print('        %-8s %9.2f' % (e.Name, float(e.Value)))
                    except Exception:
                        pass
        sev = [x for x in MSGS if 'Summary' in x or 'Errors' in x or 'Warnings' in x]
        print()
        print('=== 错误汇总 ===')
        for x in sev:
            print('   |', x[:170])
        try:
            doc.SaveAs(path.replace('.bkp', '.apwz'))
            doc.SaveAs(path.replace('.bkp', '_s.bkp'))
            print('已保存 apwz + bkp')
        except Exception as ex:
            print('保存失败:', ex)
    try:
        doc.Close()
    except Exception:
        pass


p = r'D:\<化工工作区>\_probe\p10a.bkp'
build(0.035, '10.0', 0.0050, '18.0', 0.9500, '10.0', p)
run('T-402 D:F=0.035 RR=10', p, final=True)
p = r'D:\<化工工作区>\_probe\p10b.bkp'
build(0.040, '15.0', 0.0050, '18.0', 0.9500, '10.0', p)
run('T-402 D:F=0.040 RR=15', p)
print('DONE')
