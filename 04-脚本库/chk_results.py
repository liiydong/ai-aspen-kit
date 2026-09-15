# -*- coding: utf-8 -*-
import sys, time
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as win32

DOC = r'D:\<化工工作区>\_probe\sim5_saved.bkp'
doc = win32.DispatchEx('Apwn.Document')
try:
    doc.SuppressDialogs = True
except Exception:
    pass
doc.InitFromArchive2(DOC)
time.sleep(3)

names = ['S-101', 'S-102', 'S-103', 'S-104', 'S-105', 'S-106', 'S-107', 'S-108',
         'S-109', 'S-110', 'S-111', 'S-112', 'S-113', 'S-114', 'S-115', 'S-116',
         'S-117', 'S-118', 'S-119', 'S-120', 'S-121', 'S-122', 'S-123', 'S-124',
         'S-125', 'S-201', 'S-301', 'S-401', 'S-402']
print('%-7s %10s %10s %12s %12s %8s' % ('stream', 'T(C)', 'P(bar)', 'MOLE-FLOW', 'MASS-FLOW', 'VFRAC'))
for s in names:
    nd = doc.Tree.FindNode(r'\Data\Streams\%s\Output' % s)
    if nd is None:
        print('%-7s (无 Output)' % s); continue
    vals = {}
    for f in ['TEMP', 'PRES', 'MOLE-FLOW', 'MASS-FLOW', 'VFRAC', 'MASSVFRA']:
        try:
            e = nd.Elements.Item(f)
            if e is not None:
                vals[f] = e.Value
        except Exception:
            pass
    print('%-7s %10s %10s %12s %12s %8s' % (
        s, vals.get('TEMP'), vals.get('PRES'), vals.get('MOLE-FLOW'),
        vals.get('MASS-FLOW'), vals.get('VFRAC')))

print()
print('=== T-201 输出节点 ===')
n = doc.Tree.FindNode(r'\Data\Blocks\T-201\Output')
if n is None:
    print('  无')
else:
    for f in ['TEMP', 'PRES', 'DUTY', 'B-PRES', 'B-TEMP', 'COND_DUTY', 'REB_DUTY',
              'MOLE_RR', 'VAP-FLOW', 'LIQ-FLOW']:
        try:
            e = n.Elements.Item(f)
            if e is not None:
                print('   %-12s %s' % (f, e.Value))
        except Exception:
            pass

try:
    doc.Close()
except Exception:
    pass
print('done')
