# -*- coding: utf-8 -*-
"""BIP 攻关：补 DATABANKS 库清单 + 测 ESTIMATE，读回 NRTL 二元参数条数"""
import sys, re, time, os, shutil
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

BASE = r'D:\<化工工作区>'
PROBE = os.path.join(BASE, '_probe')
SRC = os.path.join(BASE, 'NA_work.bkp')

t0 = open(SRC, encoding='utf-8', errors='ignore').read()
print('源大小', len(t0))

# ---------- 1) 补 DATABANKS ----------
DB = ('FILE-SYM-NAM = ( "APV150 PURE41" "APV150 AQUEOUS" "APV150 SOLIDS" \n'
      '"APV150 INORGANIC" "APESV150 AP-EOS" "NISTV150 NIST-TRC" ) ')

OLD = r'\ DATABANKS \ ? COMPONENTS'
print('DATABANKS 空段定位:', t0.find(OLD))
if t0.find(OLD) >= 0:
    NEW = r'\ DATABANKS ' + '\n' + DB + r'\ ? COMPONENTS'
    t1 = t0.replace(OLD, NEW, 1)
    print('1) DATABANKS 已补 ✓')
else:
    t1 = t0
    print('1) DATABANKS 未找到锚点，跳过')

# ---------- 2) 两个变体 ----------
V_NO = os.path.join(BASE, 'NA_biptest_no.bkp')
V_YES = os.path.join(BASE, 'NA_biptest_yes.bkp')
open(V_NO, 'w', encoding='utf-8', errors='ignore').write(t1)
t2, n = re.subn(r'ESTIMATE = (?:NO|YES)', 'ESTIMATE = YES', t1)
print('2) ESTIMATE->YES 替换数:', n)
open(V_YES, 'w', encoding='utf-8', errors='ignore').write(t2)

# ---------- 3) 依次打开并读参数 ----------
RES = {}


def count_pairs(doc):
    """统计 NRTL 参数条数"""
    cands = [
        r'\Data\Properties\Parameters\Binary\NRTL-1',
        r'\Data\Properties\Parameters\Binary\"T-DEPENDENT"\NRTL-1',
        r'\Data\Properties\Parameters\Binary\T-DEPENDENT\NRTL-1',
    ]
    out = {}
    for p in cands:
        try:
            nd = doc.Tree.FindNode(p)
            if nd is None:
                out[p] = 'None'
                continue
            try:
                cnt = nd.Elements.Count
            except Exception:
                cnt = '?'
            # 尝试读数据表
            try:
                val = nd.Value
            except Exception:
                val = None
            out[p] = 'Count=%s' % cnt
            # 遍历子节点
            try:
                names = [nd.Elements.Item(i).Name for i in range(min(nd.Elements.Count, 40))]
                out[p] += ' ' + str(names[:12])
            except Exception:
                pass
        except Exception as e:
            out[p] = 'ERR ' + str(e)[:60]
    return out


def run_variant(path, tag):
    print()
    print('=' * 60)
    print('变体:', tag, os.path.basename(path))
    msgs = []

    class Sink:
        def OnControlPanelMessage(self, *a):
            s = ' '.join(str(x) for x in a).strip()
            if s and s != 'False':
                msgs.append(s)

    doc = win32.DispatchEx('Apwn.Document')
    doc.SuppressDialogs = True
    win32.WithEvents(doc, Sink)
    doc.InitFromArchive2(path)
    time.sleep(3)
    doc.Engine.Run2(False)
    t = time.time()
    while time.time() - t < 420:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.25)
        if time.time() - t > 10:
            try:
                if not bool(doc.Engine.IsRunning):
                    break
            except Exception:
                break
    for _ in range(80):
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)

    # 面板统计
    err = [m for m in msgs if re.search(r'ERROR|CANNOT BE ESTIMATED|MISSING INPUT', m, re.I)]
    print('面板消息总数:', len(msgs), '| 含错误/估算失败:', len(err))
    for m in err[:6]:
        print('   !', m[:120])

    # 读参数
    r = count_pairs(doc)
    for k, v in r.items():
        print('   %-58s %s' % (k, str(v)[:110]))

    # SaveAs 保存（把内存参数落盘）
    out = os.path.join(BASE, 'NA_bipres_%s.bkp' % tag)
    try:
        doc.SaveAs(out)
        print('   SaveAs ->', os.path.basename(out), os.path.getsize(out) if os.path.exists(out) else 'NA')
    except Exception as e:
        print('   SaveAs 失败:', str(e)[:80])

    try:
        doc.Close()
    except Exception:
        pass
    return out


out_no = run_variant(V_NO, 'no')
out_yes = run_variant(V_YES, 'yes')

# ---------- 4) 文本分析保存后的文件 ----------
print()
print('=' * 60)
print('保存后文件中的 NRTL 二元对')
for tag, p in [('NO', out_no), ('YES', out_yes)]:
    if not os.path.exists(p):
        continue
    x = open(p, encoding='utf-8', errors='ignore').read()
    i = x.find('PARAMNAME = NRTL')
    print('--- ESTIMATE=%s ---' % tag)
    if i > 0:
        seg = x[i:i + 6000]
        c1 = re.findall(r'CID1\s*=\s*"?([A-Za-z0-9_-]+)"?\s*CID2\s*=\s*"?([A-Za-z0-9_-]+)"?', seg)
        print('   二元对条数:', len(c1))
        for a, b in c1[:40]:
            print('      %s - %s' % (a, b))
    else:
        print('   未找到 NRTL 段')
    # DATABANKS 是否还在
    print('   DATABANKS 含 FILE-SYM-NAM:', 'FILE-SYM-NAM' in x[i - 6000:i] if i > 6000 else 'FILE-SYM-NAM' in x)

print()
print('DONE')
