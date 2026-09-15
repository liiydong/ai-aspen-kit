# -*- coding: utf-8 -*-
"""实验：补 DATABANKS 声明，看 Aspen 是否自动检索 NRTL 二元参数"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_v13_s.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_bipA.bkp'

t = open(SRC, encoding='utf-8', errors='ignore').read()
print('原始 DATABANKS 段:')
i = t.find('? DATABANKS ?')
j = t.find('? COMPONENTS', i)
print(repr(t[i:j+20]))
print()

NEW = ('? DATABANKS ? \n\\ DATABANKS \n'
       'FILE-SYM-NAM = ( "APV150 PURE41" "APV150 AQUEOUS" "APV150 SOLIDS" '
       '"APV150 INORGANIC" "APESV150 AP-EOS" "NISTV150 NIST-TRC" ) \\ ')
t2, n = re.subn(r'\?\s*DATABANKS\s*\?.*?(?=\?\s*COMPONENTS\s+MAIN\s*\?)',
                NEW, t, count=1, flags=re.S)
print('替换次数:', n)
i = t2.find('? DATABANKS ?')
j = t2.find('? COMPONENTS', i)
print('新 DATABANKS 段:')
print(repr(t2[i:j+20]))
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t2)

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
doc.InitFromArchive2(OUT)
time.sleep(3)


def gg(p):
    n = doc.Tree.FindNode(p)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return None


# 看组分数量是否还正常
print()
print('=== 组分列表（确认没被破坏）===')
n = doc.Tree.FindNode(r'\Data\Components\Input\CID')
print('  组分节点:', n is not None, '值:', gg(r'\Data\Components\Input\CID'))

# 二元参数节点
print()
print('=== 二元参数节点探查 ===')
for p in [r'\Data\Properties\Parameters\Binary Interaction Parameters',
          r'\Data\Properties\Parameters',
          r'\Data\Properties\Parameters\Binary Interaction Parameters\NRTL-1']:
    n = doc.Tree.FindNode(p)
    print('  %-70s %s' % (p, 'OK' if n is not None else 'None'))
    if n is not None:
        try:
            print('       Elements.Count =', n.Elements.Count)
            for k in range(min(5, n.Elements.Count)):
                print('        -', n.Elements.Item(k).Name)
        except Exception as ex:
            print('       (无子节点)', ex)

doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 400:
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

print()
print('=== 面板（含 BINARY 关键信息）===')
for x in MSGS[-30:]:
    print('   |', x[:175])

print()
print('=== S-113 / S-114 组成 ===')
for s in ['S-110', 'S-113', 'S-114']:
    b = r'\Data\Streams\%s\Output' % s
    print('  --- %s  MASS=%s' % (s, gg(b + r'\MASSFLMX\MIXED')))
    n = doc.Tree.FindNode(b + r'\MASSFLOW3')
    if n is not None:
        try:
            for i2 in range(n.Elements.Count):
                e = n.Elements.Item(i2)
                try:
                    if e.Value and abs(float(e.Value)) > 0.05:
                        print('        %-8s %9.2f' % (e.Name, float(e.Value)))
                except Exception:
                    pass
        except Exception as ex:
            print('       err', ex)

try:
    doc.Close()
except Exception:
    pass
print('DONE')
