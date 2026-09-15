# -*- coding: utf-8 -*-
import sys, time, re
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_BIP2.bkp'
doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True


class S:
    def OnControlPanelMessage(self, *a):
        pass


win32.WithEvents(doc, S)
doc.InitFromArchive2(SRC)
time.sleep(2)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 400:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 8:
        try:
            if not doc.Engine.IsRunning:
                break
        except Exception:
            break
for _ in range(60):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)


def kids(path, filt=None, limit=60):
    try:
        n = doc.Tree.FindNode(path)
    except Exception as e:
        return ['ERR ' + str(e)[:50]]
    if n is None:
        return ['None']
    out = []
    try:
        c = n.Elements.Count
    except Exception:
        try:
            return ['leaf val=%s' % str(n.Value)[:40]]
        except Exception:
            return ['?']
    for i in range(min(c, limit)):
        try:
            e = n.Elements.Item(i)
            nm = e.Name
            if filt and not re.search(filt, nm, re.I):
                continue
            try:
                v = e.Value
            except Exception:
                v = ''
            out.append('%s = %s' % (nm, str(v)[:26]))
        except Exception:
            pass
    return out


print('=== S-113 / S-114 摩尔分率 ===')
for s in ['S-113', 'S-114', 'S-112']:
    try:
        nn = doc.Tree.FindNode(r'\Data\Streams\%s\Output\MOLEFRAC\MIXED' % s)
        if nn is None:
            nn = doc.Tree.FindNode(r'\Data\Streams\%s\Output\MOLEFLOW\MIXED' % s)
        d = []
        for i in range(nn.Elements.Count):
            e = nn.Elements.Item(i)
            try:
                v = float(e.Value or 0)
                if abs(v) > 1e-6:
                    d.append('%s=%.4f' % (e.Name, v))
            except Exception:
                pass
        print('  %-7s %s' % (s, ', '.join(d)))
    except Exception as ex:
        print('  %-7s ERR %s' % (s, str(ex)[:50]))

print()
print('=== T-301 (SEP) Input 子节点（含 PHASE/FLASH/TEMP/PRES/OPT）===')
for x in kids(r'\Data\Blocks\T-301\Input', r'PHASE|FLASH|TEMP|PRES|OPT|STAB|VAP|LIQ|SPEC'):
    print('   ', x)

print()
print('=== Setup Input 子节点 ===')
for x in kids(r'\Data\Setup\Input'):
    print('   ', x)

print()
print('=== Components\\Henry-Comps\\Input 子节点 ===')
for x in kids(r'\Data\Components\Henry-Comps\Input'):
    print('   ', x)

print()
print('=== T-202 (RADFRAC) Input 子节点（含 PHASE/FLASH/OPT）===')
for x in kids(r'\Data\Blocks\T-202\Input', r'PHASE|FLASH|OPT|STAB|CONV|NSTAGE|CHEM'):
    print('   ', x)

doc.SaveAs(r'D:\<化工工作区>\_probe\na_probe_save.bkp')
try:
    doc.Close()
except Exception:
    pass
print('DONE')
