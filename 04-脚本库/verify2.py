# -*- coding: utf-8 -*-
"""重跑后读塔节点（正确名）+ 查组分沸点"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

SRC = r'D:\<化工工作区>\NA-Chemical-10000t_最终定稿.bkp'
doc = win32.DispatchEx('Apwn.Document')
doc.SuppressDialogs = True
doc.InitFromArchive2(SRC)
time.sleep(3)
doc.Engine.Run2(False)
t0 = time.time()
while time.time() - t0 < 900:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.25)
    if time.time() - t0 > 10:
        try:
            if not bool(doc.Engine.IsRunning):
                break
        except Exception:
            break
for _ in range(100):
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

def names(path):
    n = doc.Tree.FindNode(path)
    if n is None:
        return []
    out = []
    try:
        for i in range(n.Elements.Count):
            out.append(n.Elements.Item(i).Name)
    except Exception:
        pass
    return out

L = []
# 1) 组分身份 + 沸点
L.append('===== Component 节点 =====')
cn = doc.Tree.FindNode(r'\Data\Components')
cids = [cn.Elements.Item(i).Name for i in range(cn.Elements.Count)] if cn is not None else []
L.append('  组分: %s' % ', '.join(cids))
if cids:
    sub = names(r'\Data\Components\%s' % cids[0])
    L.append('  %s 子节点样例: %s' % (cids[0], ' | '.join(sub[:40])))
    L.append('')
    for c in cids:
        vals = {}
        for k in ['ALIAS', 'TB', 'TC', 'MW', 'CASN', 'CASNO', 'PC']:
            v = g(r'\Data\Components\%s\%s' % (c, k))
            if v is not None:
                vals[k] = str(v)[:24]
        L.append('  %-10s %s' % (c, vals))

# 2) 塔节点
L.append('')
L.append('===== 塔节点实测 =====')
CAND_TOP = ['TOPD_TEMP', 'TOPD_VFLOW', 'TOPD_LFLOW', 'TOP_TEMP', 'TOP_PRES',
            'BOTM_TEMP', 'BOTM_VFLOW', 'BOTM_LFLOW', 'BOTTOM_TEMP', 'BOT_LFLOW', 'BOT_VFLOW',
            'CD_TEMP', 'REB_TEMP', 'COND_TEMP', 'REB_DUTY', 'COND_DUTY', 'RR', 'DF', 'D_F',
            'DIAM', 'CA_DIAM1', 'CA_DIAM2', 'DIAM2', 'ACT_AREA', 'FEED_VFLOW', 'FEED_LFLOW',
            'FEED_FLOW', 'VAP_VFLOW', 'LIQ_VFLOW', 'AVGDP_HT1', 'HT_FROM_TOP']
for b in ['T-201', 'T-302', 'T-401', 'T-402', 'T-404']:
    L.append('  ---- %s' % b)
    for k in CAND_TOP:
        v = g(r'\Data\Blocks\%s\Output\%s' % (b, k))
        if v is not None:
            L.append('     %-14s = %s' % (k, str(v)[:48]))
    # 带 TOP 名字的输出
    for k in names(r'\Data\Blocks\%s\Output' % b):
        if 'TOP' in k and k not in CAND_TOP:
            v = g(r'\Data\Blocks\%s\Output\%s' % (b, k))
            if v is not None and str(v) not in ('0', '0.0', 'None'):
                L.append('     [t] %-14s = %s' % (k, str(v)[:48]))

open(r'D:\<化工工作区>\_probe\nodes2.txt', 'w', encoding='utf-8').write('\n'.join(L))
try:
    doc.Close()
except Exception:
    pass
print('DONE')
