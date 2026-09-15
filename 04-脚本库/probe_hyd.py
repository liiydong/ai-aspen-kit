# -*- coding: utf-8 -*-
import sys, time, re
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_BIP.bkp'
MSGS = []


class Sink:
    def OnControlPanelMessage(self, *a):
        pass


doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
win32.WithEvents(doc, Sink)
doc.InitFromArchive2(SRC)
time.sleep(2)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 400:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 8:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(60):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)

for bid in ['T-401', 'T-404']:
    for base in [r'\Data\Blocks\%s\Output' % bid, r'\Data\Blocks\%s\Hydraulics' % bid]:
        try:
            n = doc.Tree.FindNode(base)
        except Exception:
            n = None
        print('==', base, '->', 'None' if n is None else 'ok')
        if n is None:
            continue
        try:
            c = n.Elements.Count
        except Exception:
            continue
        for i in range(min(c, 60)):
            try:
                e = n.Elements.Item(i)
                try:
                    v = e.Value
                except Exception:
                    v = ''
                try:
                    sub = e.Elements.Count
                except Exception:
                    sub = '-'
                print('   %-34s val=%-16s children=%s' % (e.Name, str(v)[:16], sub))
            except Exception:
                pass

# 常见节点直接试
print()
print('--- 候选节点 ---')
for p in [r'\Data\Blocks\T-401\Output\TOP_TEMP', r'\Data\Blocks\T-401\Output\BOTTOM_TEMP',
          r'\Data\Blocks\T-401\Output\TOP_VFLOW', r'\Data\Blocks\T-401\Output\BOT_VFLOW',
          r'\Data\Blocks\T-401\Output\TOP_LFLOW', r'\Data\Blocks\T-401\Output\BOT_LFLOW',
          r'\Data\Blocks\T-401\Output\RR', r'\Data\Blocks\T-401\Output\MOLE_D',
          r'\Data\Blocks\T-401\Output\MASS_D', r'\Data\Blocks\T-401\Output\COND_DUTY',
          r'\Data\Blocks\T-401\Output\REB_DUTY']:
    try:
        n = doc.Tree.FindNode(p)
        print('  %-56s = %s' % (p.split('Output')[-1], (n.Value if n is not None else None)))
    except Exception as e:
        print('  %-56s ERR %s' % (p.split('Output')[-1], str(e)[:40]))
doc.Close()
doc.Quit()
print('DONE')
