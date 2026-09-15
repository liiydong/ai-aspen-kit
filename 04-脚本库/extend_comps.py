# -*- coding: utf-8 -*-
"""验证：在 bkp 文本的 COMPONENTS 段按真实格式追加组分，Aspen 是否接受"""
import time, sys, re, shutil
import win32com.client as win32

SRC = r'D:\<化工工作区>\_probe\x_NA-Chemical-10000t.bkp.backup'
BAS = r'D:\<化工工作区>\_probe\base.bkp'
MOD = r'D:\<化工工作区>\_probe\mod.bkp'

shutil.copyfile(SRC, BAS)
text = open(SRC, encoding='utf-8', errors='ignore').read()

# ---- 定位 COMPONENTS 段 ----
start = text.find('? COMPONENTS MAIN ?')
marker = '\\ ? COMPONENTS "ADA/PCS"'
end = text.find(marker, start)
print('锚点 start=%d  end=%d' % (start, end))
print('现有段预览:', repr(text[start:start+120]))
print('段尾预览  :', repr(text[end-80:end+20]))
print()

# ---- 现有组分 ----
seg = text[start:end]
cids = re.findall(r'CID = ("?[^"\s/]+"?)', seg)
print('现有组分 (%d):' % len(cids), [c.strip('"') for c in cids])
print()

# ---- 要追加的 6 个 ----
NEW = [
    ('4-MP',   '4-METHYLPYRIDINE', 'C6H7N-N2'),
    ('NAM',    'NICOTINAMIDE',     'C6H6N2O'),
    ('NAC',    'NICOTINIC-ACID',   'C6H5NO2'),
    ('NAOH',   'SODIUM-HYDROXIDE', 'HNAO'),
    ('H2SO4',  'SULFURIC-ACID',    'H2O4S'),
    ('NA2SO4', 'SODIUM-SULFATE',   'NA2O4S'),
]

def entry(cid, db, al):
    return ('CID = "%s" ANAME = %s OUTNAME = "%s" TYPE = CONV DBNAME1 = "%s" ANAME1 = "%s"'
            % (cid, al, cid, db, al))

add = ' /  ' + ' /  '.join(entry(*c) for c in NEW) + ' '
newtext = text[:end] + add + text[end:]
open(MOD, 'w', encoding='utf-8', errors='ignore').write(newtext)
print('已写出 mod.bkp，长度 %d -> %d' % (len(text), len(newtext)))
print()


def probe(path, tag):
    print('=' * 18, tag, path)
    doc = win32.DispatchEx('Apwn.Document')
    ok = False
    for m in ('InitFromArchive2', 'InitFromArchive'):
        try:
            getattr(doc, m)(path)
            print('  打开方式:', m)
            ok = True
            break
        except Exception as e:
            print('  %s 失败: %s' % (m, str(e)[:100]))
    if not ok:
        return
    time.sleep(3)
    node = doc.Tree.FindNode(r'\Data\Components\Specifications\Input')
    n = node.Elements.Count
    print('  Input 子节点数:', n)
    names = []
    for i in range(n):
        try:
            nm = node.Elements.Item(i).Name
            names.append(nm)
        except Exception:
            pass
    print('  子节点:', names)
    # 组分实例
    try:
        comp = doc.Tree.FindNode(r'\Data\Components\Main')
        print('  Main 实例数:', comp.Elements.Count)
        inst = [comp.Elements.Item(i).Name for i in range(comp.Elements.Count)]
        print('  组分实例:', inst)
    except Exception as e:
        print('  Main 读取失败:', str(e)[:100])
    try:
        doc.Close()
    except Exception:
        pass
    print()


probe(BAS, '[基线] 未修改')
probe(MOD, '[修改] 追加6个')
