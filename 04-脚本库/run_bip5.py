# -*- coding: utf-8 -*-
"""运行 BDBANK 版并回存，解析 Aspen 实际检索到的 NRTL 二元参数"""
import sys, re, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32
import pythoncom

P = r'D:\<化工工作区>\_probe\bipA2.bkp'
SAVE = r'D:\<化工工作区>\_probe\bipA2_s.bkp'

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive2(P)
time.sleep(3)
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
for _ in range(50):
    pythoncom.PumpWaitingMessages()
    time.sleep(0.05)
try:
    doc.SaveAs(SAVE)
    print('已回存:', SAVE)
except Exception as ex:
    print('回存失败:', ex)
try:
    doc.Close()
except Exception:
    pass

# 解析 NRTL-1 段
t = open(SAVE, encoding='utf-8', errors='ignore').read()
i = t.find('"NRTL-1"')
j = t.find('? PROPERTIES', i + 10)
blk = t[i:j if j > 0 else i + 40000]
print('NRTL-1 段长度:', len(blk))
recs = re.findall(r'CID1 = "?([A-Za-z0-9-]+)"?\s+CID2 = "?([A-Za-z0-9-]+)"?(.{0,1400}?)(?=/ |\\ )', blk, re.S)
print('记录数:', len(recs))
print()
print('%-8s %-8s %-14s %-14s %-10s %s' % ('CID1', 'CID2', 'uval1', 'uval2', 'uval5', '来源标签'))
for c1, c2, body in recs:
    def g(n):
        m = re.search(r'UVAL%d = (-?[\d.eE+-]+)' % n, body)
        return m.group(1) if m else '-'
    src = re.search(r'VAL1 = "([^"]+)"', body)
    src = src.group(1) if src else '-'
    print('%-8s %-8s %-14s %-14s %-10s %s' % (c1, c2, g(1), g(2), g(5), src))
