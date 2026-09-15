# -*- coding: utf-8 -*-
import sys, time, os
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32

cands = [r'D:\<化工工作区>\_probe\dg_Bhot_saved.bkp',
         r'D:\<化工工作区>\_probe\dg_Aswap_saved.bkp']
DOC = next((p for p in cands if os.path.exists(p)), None)
print('DOC =', DOC)

doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive2(DOC)
time.sleep(3)


def g(path):
    n = doc.Tree.FindNode(path)
    if n is None:
        return None
    try:
        return n.Value
    except Exception:
        return 'ERR'


names = ['S-101', 'S-102', 'S-103', 'S-104', 'S-105', 'S-106', 'S-107', 'S-108',
         'S-109', 'S-110', 'S-112', 'S-114', 'S-116']
print('%-7s %12s %12s %14s %8s %10s' % ('stream', 'T(C)', 'P(bar)', 'MOLE-FLOW', 'VFRAC', 'MASS-FLOW'))
for s in names:
    b = r'\Data\Streams\%s\Output' % s
    print('%-7s %12s %12s %14s %8s %10s' % (
        s, g(b + r'\TEMP'), g(b + r'\PRES'), g(b + r'\MOLE-FLOW'),
        g(b + r'\VFRAC'), g(b + r'\MASS-FLOW')))

print()
print('组分流量 S-108:')
mx = g(r'\Data\Streams\S-108\Output\MOLE-FLOW')
print('  MOLE-FLOW =', mx)
for c in ['3-MP', 'NH3', 'O2', 'N2', 'H2O', '3-CP', 'CO2', 'HCN']:
    v = g(r'\Data\Streams\S-108\Output\MOLE-FLOW\%s' % c)
    if v is not None:
        print('   %-8s %s' % (c, v))

try:
    doc.Close()
except Exception:
    pass
print('done')
