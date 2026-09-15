# -*- coding: utf-8 -*-
"""切换到 SI 单位制后重跑，读真实负荷单位"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_最终定稿.bkp'
OUT = r'D:\<化工工作区>\NA-Chemical-10000t_SI.bkp'
L = []

# 文本方式切 SI（IN-UNITS INSET）
t = open(SRC, encoding='utf-8', errors='ignore').read()
t2, n = re.subn(r'INSET\s*=\s*METCBAR', 'INSET = SI', t)
L.append('文本 INSET 替换 %d 处' % n)
if n == 0:
    L.append('未找到 INSET = METCBAR')
open(OUT, 'w', encoding='utf-8', errors='ignore').write(t2)

MSGS = []
class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            MSGS.append(s)

doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
win32.WithEvents(doc, Sink)
doc.InitFromArchive2(OUT)
time.sleep(3)

# 也尝试用 COM 设 SI
for p in [r'\Data\Setup\Global\Input\INSET', r'\Data\Setup\Main\Input\INSET']:
    nd = doc.Tree.FindNode(p)
    if nd is not None:
        L.append('节点存在: %s = %s' % (p, nd.Value))
        try:
            nd.Value = 'SI'
            L.append('  已设为 SI')
        except Exception as ex:
            L.append('  设置失败 %s' % ex)

doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 600:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 10:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(80):
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

def us(p):
    n = doc.Tree.FindNode(p)
    if n is None:
        return None
    for attr in ['UnitString', 'Unit String', 'Unit']:
        try:
            v = n.AttributeValue(attr)
            if v:
                return v
        except Exception:
            pass
    return None

L.append('')
L.append('===== 面板尾 =====')
for x in MSGS[-12:]:
    L.append('  ' + x[:150])

L.append('')
L.append('===== 塔负荷（SI）=====')
for b in ['T-201', 'T-202', 'T-302', 'T-401', 'T-402', 'T-404']:
    L.append('  %-7s COND=%-14s %-12s REB=%-14s %-12s RR=%s TOP=%s BOT=%s' % (
        b,
        g(r'\Data\Blocks\%s\Output\COND_DUTY' % b), us(r'\Data\Blocks\%s\Output\COND_DUTY' % b),
        g(r'\Data\Blocks\%s\Output\REB_DUTY' % b), us(r'\Data\Blocks\%s\Output\REB_DUTY' % b),
        g(r'\Data\Blocks\%s\Output\RR' % b),
        g(r'\Data\Blocks\%s\Output\TOP_TEMP' % b),
        g(r'\Data\Blocks\%s\Output\BOTTOM_TEMP' % b)))

L.append('')
L.append('===== 单位制 =====')
for p in [r'\Data\Setup\Global\Input\INSET', r'\Data\Setup\Global\Input\UNITSET',
          r'\Data\Setup\Main\Input\INSET']:
    L.append('  %s = %s' % (p, g(p)))

open(r'D:\<化工工作区>\_probe\si_units.txt', 'w', encoding='utf-8').write('\n'.join(L))
try:
    doc.SaveAs(OUT.replace('.bkp', '_s.bkp'))
except Exception as ex:
    L.append('save fail %s' % ex)
try:
    doc.Close()
except Exception:
    pass
print('DONE')
