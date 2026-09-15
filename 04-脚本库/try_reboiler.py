# -*- coding: utf-8 -*-
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_sim6.bkp'
T0 = open(BASE, encoding='utf-8', errors='ignore').read()


def load(path, tag):
    doc = win32.DispatchEx('Apwn.Document')
    try:
        doc.SuppressDialogs = True
    except Exception:
        pass
    doc.InitFromArchive2(path)
    time.sleep(3)
    return doc


# ---------- 1) 列出 T-201 Input 里与进料/规格相关的字段 ----------
doc = load(BASE, 'dump')
n = doc.Tree.FindNode(r'\Data\Blocks\T-201\Input')
names = []
for i in range(n.Elements.Count):
    try:
        e = n.Elements.Item(i)
        nm = e.Name
        try:
            cnt = e.Elements.Count
        except Exception:
            cnt = -1
        names.append((nm, cnt))
    except Exception:
        pass
print('T-201 Input 字段中与进料/规格相关：')
for nm, cnt in names:
    if any(k in (nm or '').upper() for k in ['FEED', 'STAGE', 'PROD', 'SPEC', 'RDV', 'RR', 'BR',
                                             'PUMP', 'L1', 'VN', 'QN', 'Q1', 'DUTY']):
        print('   %-22s [%d]' % (nm, cnt))
print('   (共 %d 个字段)' % len(names))
try:
    doc.Close()
except Exception:
    pass
time.sleep(1)

# ---------- 2) 给 T-201 加再沸器（小汽化率），看能否消除干板 ----------
OLD = re.search(r'\\ \\ "COL-CONFIG" CONDENSER = NONE REBOILER = NONE ', T0)
print('COL-CONFIG 匹配:', bool(OLD))
if OLD:
    T1 = T0[:OLD.start()] + '\\ \\ "COL-CONFIG" CONDENSER = NONE REBOILER = KETTLE ' + T0[OLD.end():]
    cs = re.search(r'\\ \\ "COL-SPECS" BASIS-RDV = 1\.0 <0> <0> ', T1)
    print('COL-SPECS 匹配:', bool(cs))
    if cs:
        T1 = (T1[:cs.start()] +
              '\\ \\ "COL-SPECS" BASIS-RDV = 1.0 <0> <0> BASIS-BR = 0.1 <-1> <0> ' +
              T1[cs.end():])
        out = r'D:\<化工工作区>\_probe\reb.bkp'
        open(out, 'w', encoding='utf-8', errors='ignore').write(T1)
        print('已写:', out)
        msgs = []

        class Sink:
            def OnControlPanelMessage(self, *a):
                s = ' '.join(str(x) for x in a).strip()
                if s and s != 'False':
                    msgs.append(s)

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
        t1 = time.time()
        while time.time() - t1 < 300:
            pythoncom.PumpWaitingMessages()
            time.sleep(0.25)
            if time.time() - t1 > 8:
                try:
                    if not bool(doc.Engine.IsRunning):
                        break
                except Exception:
                    break
        for _ in range(50):
            pythoncom.PumpWaitingMessages()
            time.sleep(0.05)
        print('面板 %d 条' % len(msgs))
        for x in msgs:
            if any(k in x for k in ['ERROR', 'T-201', 'SEVERE', 'stopped', 'Converged', 'converged']):
                print('   |', x[:190])
        try:
            doc.SaveAs(r'D:\<化工工作区>\_probe\reb_saved.bkp')
        except Exception:
            pass
        try:
            doc.Close()
        except Exception:
            pass
print('DONE')
