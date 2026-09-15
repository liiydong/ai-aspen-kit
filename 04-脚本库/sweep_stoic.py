# -*- coding: utf-8 -*-
"""RStoic 多行 STOIC 表格式扫描：用控制面板报错当判据"""
import sys, re, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>\NA-Chemical-10000t_模拟.bkp'
TMP = r'D:\<化工工作区>\_probe\sweep_%s.bkp'

t0 = open(BASE, encoding='utf-8', errors='ignore').read()

# R-101 的 5 个组分（反应 1）
SPEC1 = [('3-MP', -1.0), ('NH3', -1.0), ('O2', -1.5), ('3-CP', 1.0), ('H2O', 3.0)]


def make(variant, spec, quote=True):
    def cid(c):
        return '"%s"' % c if quote else c
    rows = []
    for i, (c, co) in enumerate(spec):
        if variant == 'A':          # 全部裸名
            s = ''
        elif variant == 'B':        # 后缀 = i
            s = '' if i == 0 else str(i)
        elif variant == 'D':        # 全部裸名 + 单行
            s = ''
        else:                       # C：前两条裸，之后 1,2,...
            s = '' if i < 2 else str(i - 1)
        rows.append('\\ \\ STOIC%s REACNO%s = 1 STOIC-CID%s = %s STOIC-SSID%s = MIXED '
                    'COEF%s = %s <0> <0> ' % (s, s, s, cid(c), s, s, co))
    body = ''.join(rows)
    if variant == 'D':
        body = body.replace('\n', '')
    else:
        body = '\n'.join(rows)
    return ('? BLOCK RSTOIC "R-101" ? ; "METCBAR_MOLE" ; ; ICON1 ; \n'
            '\\ PARAM TEMP = 405.0 <22> <4> PRES = 1.8 <20> <5> SPEC-OPT = TP \n'
            + body +
            '\\ \\ CONVEX EXT-REACNO = 1 KEY-SSID = MIXED KEY-CID = %s CONV = .86 <0> <0> \n'
            '\\ \\ PRODUCTS SID = "S-106" \\ \n' % cid('3-MP'))


def patch(text, newsec):
    m = re.search(r'\?\s*BLOCK\s+RSTOIC\s+"?R-101\s*"?\s*\?', text)
    if not m:
        return None
    nxt = re.search(r'\n\?\s*BLOCK\s', text[m.start() + 10:])
    end = m.start() + 10 + nxt.start() if nxt else len(text)
    return text[:m.start()] + newsec + text[end:]


def run_one(variant, quote=True):
    out = TMP % variant
    nt = patch(t0, make(variant, SPEC1, quote))
    if nt is None:
        print('  无法定位 R-101'); return
    open(out, 'w', encoding='utf-8', errors='ignore').write(nt)

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
    while time.time() - t1 < 120:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.2)
        if time.time() - t1 > 8:
            try:
                if not bool(doc.Engine.IsRunning):
                    break
            except Exception:
                break
    for _ in range(30):
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)

    rel = [m for m in msgs if 'R-101' in m or 'MASS BALANCE' in m or 'ABSOLUTE ERROR' in m]
    print('  变体 %s (quote=%s) -> 面板 %d 条, 相关 %d 条' % (variant, quote, len(msgs), len(rel)))
    for m in rel[:8]:
        print('     |', m[:200])
    if not rel:
        print('     ✓ 无 R-101 报错')
    # 读回 COEF 行数
    for node in ['COEF', 'COEF1', 'COEF2']:
        n = doc.Tree.FindNode(r'\Data\Blocks\R-101\Input\%s' % node)
        if n is not None:
            try:
                print('     %s rows=%d' % (node, n.Elements.Count))
            except Exception:
                pass
    try:
        doc.Close()
    except Exception:
        pass
    time.sleep(1)


for v in ['A', 'B', 'C', 'D']:
    try:
        run_one(v)
    except Exception as ex:
        print('  变体 %s 异常: %s' % (v, ex))
    print()

print('ALL DONE')
