# -*- coding: utf-8 -*-
"""对照实验：Aspen 自带示例（结构/参数状态不同）下，设 ESTIMATE=YES 能否估算出参数"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

FILES = [
    (r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\Bulk Chemicals\pfdtut.bkp', 'pfdtut 5组分'),
    (r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\Bulk Chemicals\Distillation\3phase.bkp', '3phase 4组分'),
]

NRT = r'\Data\Properties\Parameters\Binary Interaction\NRTL-1'
EST = r'\Data\Properties\Estimation\Estimate\Input'


def count_pairs(doc):
    """数 NRTL 参数行（Output\CID1 的行数，Input 里 VAL1 的行数也看）"""
    r = {}
    for p, k in [(NRT + r'\Output\CID1', 'OutCID1'),
                 (NRT + r'\Output\VALUE', 'OutVALUE'),
                 (NRT + r'\Input\VAL1', 'InVAL1'),
                 (NRT + r'\Input\BDBANK', 'BDBANK')]:
        nd = doc.Tree.FindNode(p)
        if nd is None:
            r[k] = 'None'
            continue
        try:
            r[k] = nd.Elements.Count
        except Exception:
            r[k] = '?'
    return r


def probe(path, tag):
    msgs = []

    class Sink:
        def OnControlPanelMessage(self, *a):
            s = ' '.join(str(x) for x in a).strip()
            if s and s != 'False':
                msgs.append(s)

    print()
    print('=' * 70)
    print(tag)
    doc = win32.DispatchEx('Apwn.Document')
    doc.SuppressDialogs = True
    win32.WithEvents(doc, Sink)
    doc.InitFromArchive2(path)
    t = time.time()
    while time.time() - t < 20:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.25)

    print('  打开后:', count_pairs(doc))
    # 结构状态
    for c in list({'_'}.union()):
        pass
    nd = doc.Tree.FindNode(r'\Data\Components\Specifications\Input\ANAME')
    comps = []
    if nd is not None:
        try:
            for i in range(nd.Elements.Count):
                comps.append(nd.Elements.Item(i).Name)
        except Exception:
            pass
    print('  组分:', comps[:8])
    if comps:
        c0 = comps[0]
        for col in ['ATOMTYPE', 'BONDTYPE']:
            n2 = doc.Tree.FindNode(r'\Data\Properties\Molecular Structure\%s\Input\%s' % (c0, col))
            try:
                print('    %s / %s 行数 = %s' % (c0, col, n2.Elements.Count))
            except Exception:
                print('    %s / %s -> ?' % (c0, col))

    # 设估算开关
    for p, v in [(EST + r'\ALLONLY', 'ALL'), (NRT + r'\Input\ESTIMATE', 'YES')]:
        nd = doc.Tree.FindNode(p)
        try:
            nd.Elements.Item(0).Value = v
        except Exception:
            try:
                nd.Value = v
            except Exception as e:
                print('  设 %s 失败: %s' % (p.split('\\')[-1], str(e)[:60]))

    n0 = len(msgs)
    doc.Engine.Run2(False)
    t = time.time()
    while time.time() - t < 240:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.3)
        if time.time() - t > 8:
            try:
                if not bool(doc.Engine.IsRunning):
                    break
            except Exception:
                break
    for _ in range(80):
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)
    new = msgs[n0:]
    est = [m for m in new if re.search(r'ESTIMAT|UNIFAC|PCES', m)]
    print('  运行后:', count_pairs(doc))
    print('  估算相关消息 %d 条:' % len(est))
    for m in est[:6]:
        print('     |', m[:120])
    try:
        doc.Close()
    except Exception:
        pass


for p, tag in FILES:
    if os.path.exists(p):
        try:
            probe(p, tag)
        except Exception as e:
            print('  ERR:', str(e)[:120])
    else:
        print('缺失', p)
print('DONE')
