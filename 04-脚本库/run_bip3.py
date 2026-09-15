# -*- coding: utf-8 -*-
"""实验：给 NRTL 参数段补 BDBANK 检索清单，让 Aspen 自动检索二元参数"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_v13_s.bkp'
base = open(SRC, encoding='utf-8', errors='ignore').read()

BDBANK = (' BDBANK = ( "APV150 VLE-RK" "APV150 VLE-IG" "APV150 LLE-ASPEN" '
          '"APV150 VLE-LIT" "APV150 VLE-HOC" ) ESTIMATE = NO ')

# 变体 A：只给 NRTL 补 BDBANK
tA, nA = re.subn(r'(PROP-LIST PARAMNAME = NRTL SETNO = 1 UNITROW = 0 TUNITROW = 22 TUNITLABEL = F)( NEL = 12)',
                 lambda m: m.group(1) + BDBANK + 'NEL = 12', base, count=1)
print('A 替换:', nA)

# 变体 B：BDBANK + DATABANKS 声明
tB, nB = re.subn(r'(PROP-LIST PARAMNAME = NRTL SETNO = 1 UNITROW = 0 TUNITROW = 22 TUNITLABEL = F)( NEL = 12)',
                 lambda m: m.group(1) + BDBANK + 'NEL = 12', base, count=1)
DBLIST = ('? DATABANKS ? \n\\ DATABANKS \nFILE-SYM-NAM = ( "APV150 PURE41" "APV150 AQUEOUS" '
          '"APV150 SOLIDS" "APV150 INORGANIC" "APESV150 AP-EOS" "NISTV150 NIST-TRC" '
          '"APV150 VLE-RK" "APV150 VLE-IG" "APV150 LLE-ASPEN" "APV150 VLE-LIT" "APV150 VLE-HOC" ) \\ ')
tB, n2 = re.subn(r'\?\s*DATABANKS\s*\?.*?(?=\?\s*COMPONENTS\s+MAIN\s*\?)', DBLIST, tB, count=1, flags=re.S)
print('B DATABANKS 替换:', n2)

open(r'D:\<化工工作区>\_probe\bipA.bkp', 'w', encoding='utf-8', errors='ignore').write(tA)
open(r'D:\<化工工作区>\_probe\bipB.bkp', 'w', encoding='utf-8', errors='ignore').write(tB)


def run(tag, path):
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
    doc.InitFromArchive2(path)
    time.sleep(3)

    def g(p):
        n = doc.Tree.FindNode(p)
        if n is None:
            return None
        try:
            return n.Value
        except Exception:
            return 'ERR'

    # 先看二元参数节点
    bin_paths = [r'\Data\Properties\Parameters\Binary Interaction',
                 r'\Data\Properties\Parameters\Binary Interaction\NRTL-1']
    print('=' * 72)
    print('变体', tag)
    nd = doc.Tree.FindNode(bin_paths[0])
    if nd is not None:
        try:
            print('  Binary Interaction 子节点 %d 个:' % nd.Elements.Count)
            for k in range(nd.Elements.Count):
                print('    -', nd.Elements.Item(k).Name)
        except Exception as ex:
            print('  ', ex)
    n1 = doc.Tree.FindNode(bin_paths[1])
    if n1 is not None:
        try:
            print('  NRTL-1 记录数:', n1.Elements.Count)
        except Exception as ex:
            print('  NRTL-1:', ex)

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

    errs = [x for x in MSGS if re.search(r'ERROR|MISSING|BINARY|INVALID|stopped', x, re.I)]
    print('  关键消息:')
    for x in errs[:8]:
        print('    |', x[:165])
    print('  结果:')
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


run('A: 仅 BDBANK', r'D:\<化工工作区>\_probe\bipA.bkp')
run('B: BDBANK + DATABANKS', r'D:\<化工工作区>\_probe\bipB.bkp')
print('DONE')
