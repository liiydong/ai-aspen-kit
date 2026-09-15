# -*- coding: utf-8 -*-
"""实验：正确插入 BDBANK 检索清单；并注入水-甲苯 NRTL 实测参数验证机制"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_v13_s.bkp'
base = open(SRC, encoding='utf-8', errors='ignore').read()


def locate_nrtl_end(t):
    """返回 NRTL-1 段中 NEL = 12 之后那个分隔反斜杠的位置"""
    i = t.find('"NRTL-1"')
    if i < 0:
        return -1
    j = t.find('NEL = 12', i)
    if j < 0:
        return -1
    k = t.find('\\', j)          # 段分隔反斜杠
    return k


BDBANK = ('BDBANK = ( "APV150 VLE-RK" "APV150 VLE-IG" "APV150 LLE-ASPEN" '
          '"APV150 VLE-LIT" "APV150 VLE-HOC" ) ESTIMATE = NO ')

# ---- 变体 A2：只插 BDBANK ----
k = locate_nrtl_end(base)
tA = base[:k] + '\\ \\ BPVAL PARAMNAME2 = NRTL CID1 = H2O CID2 = TOL UNITROW2 = 0 TUNITROW2 = 22 TUNITLABEL2 = C ' + \
     'VAL1 = "APV150 LLE-ASPEN" VAL2 = "APV150 LLE-ASPEN" VAL3 = "APV150 LLE-ASPEN" VAL4 = "APV150 LLE-ASPEN" ' + \
     'VAL5 = "APV150 LLE-ASPEN" VAL6 = "APV150 LLE-ASPEN" VAL7 = "APV150 LLE-ASPEN" VAL8 = "APV150 LLE-ASPEN" ' + \
     'VAL9 = "APV150 LLE-ASPEN" VAL10 = "APV150 LLE-ASPEN" VAL11 = "APV150 LLE-ASPEN" VAL12 = "APV150 LLE-ASPEN" / ' + \
     base[k:]
# 在 PROP-LIST 里加 BDBANK
tA2, nA = re.subn(r'(PARAMNAME = NRTL SETNO = 1 UNITROW = 0 TUNITROW = 22 TUNITLABEL = F)(\s+)(NEL = 12)',
                  lambda m: m.group(1) + ' ' + BDBANK + m.group(3), tA, count=1)
print('A2 BDBANK 插入:', nA)

# ---- 变体 D：BDBANK + 水-甲苯数值实测（来自 Aspen 自带 Conceptual 示例） ----
TOLREC = ('BPVAL PARAMNAME2 = NRTL CID1 = H2O CID2 = TOL UNITROW2 = 0 TUNITROW2 = 22 TUNITLABEL2 = C '
          'UVAL1 = -5.05265 <0> <0> UVAL2 = 5.01128 <0> <0> UVAL3 = 2182.52 <0> <0> UVAL4 = 473.976 <0> <0> '
          'UVAL5 = 0.100364 <0> <0> UVAL11 = -3.497 <0> <0> UVAL12 = 250.071 <0> <0> '
          'VAL1 = "0$-5.05265" VAL2 = "0$5.01128" VAL3 = "0$2182.52" VAL4 = "0$473.976" VAL5 = "0$0.100364" '
          'VAL6 = "0$0.0" VAL7 = "0$0.0" VAL8 = "0$0.0" VAL9 = "0$0.0" VAL10 = "0$0.0" '
          'VAL11 = "0$-3.497" VAL12 = "0$250.071" / ')
k = locate_nrtl_end(base)
tD = base[:k] + '\\ \\ ' + TOLREC + base[k:]
tD, nD = re.subn(r'(PARAMNAME = NRTL SETNO = 1 UNITROW = 0 TUNITROW = 22 TUNITLABEL = F)(\s+)(NEL = 12)',
                 lambda m: m.group(1) + ' ' + BDBANK + m.group(3), tD, count=1)
print('D 插入:', nD)

open(r'D:\<化工工作区>\_probe\bipA2.bkp', 'w', encoding='utf-8', errors='ignore').write(tA2)
open(r'D:\<化工工作区>\_probe\bipD.bkp', 'w', encoding='utf-8', errors='ignore').write(tD)


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

    print('=' * 72)
    print('变体', tag)
    nd = doc.Tree.FindNode(r'\Data\Properties\Parameters\Binary Interaction\NRTL-1')
    if nd is not None:
        try:
            print('  NRTL-1 子记录数:', nd.Elements.Count)
            for k2 in range(min(6, nd.Elements.Count)):
                print('    -', nd.Elements.Item(k2).Name)
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

    for x in MSGS:
        if re.search(r'BINARY|MISSING|ERROR|ZERO|stopped', x, re.I):
            print('   |', x[:165])
    print('  结果:')
    for s in ['S-112', 'S-113', 'S-114']:
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


run('A2: BDBANK(空值) ', r'D:\<化工工作区>\_probe\bipA2.bkp')
run('D: BDBANK + 水甲苯实测值', r'D:\<化工工作区>\_probe\bipD.bkp')
print('DONE')
