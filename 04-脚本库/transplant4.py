# -*- coding: utf-8 -*-
"""移植 v4：不再依赖组分名正则，直接按 MOLEC-STRUCT 出现位置整块搬运"""
import sys, re, os
sys.stdout.reconfigure(encoding='utf-8')

BASE = r'D:\<化工工作区>'
USER = os.path.join(BASE, 'NA_user.bkp')
FULL = os.path.join(BASE, 'NA_work.bkp')
OUT = os.path.join(BASE, 'NA_struct.bkp')

U = open(USER, encoding='utf-8', errors='ignore').read()
W = open(FULL, encoding='utf-8', errors='ignore').read()


def struct_block(x, tag):
    """返回 MOLEC-STRUCT 结构区块 [start, end)"""
    idx = [m.start() for m in re.finditer(r'MOLEC-STRUCT', x)]
    print('%s: MOLEC-STRUCT 出现 %d 次' % (tag, len(idx)))
    if not idx:
        return None, None
    # 起点：往前找最近的 '? PROPERTIES'
    s = x.rfind('? PROPERTIES', 0, idx[0])
    if s < 0:
        s = idx[0]
    # 终点：最后一个出现之后，找下一个 '? PROPERTIES'
    e = x.find('? PROPERTIES', idx[-1])
    if e < 0:
        e = len(x)
    return s, e


su = struct_block(U, '用户文件')
sw = struct_block(W, '完整模型')
print('  用户区块: %s  长度=%d' % (su, (su[1] - su[0]) if su[0] is not None else 0))
print('  模型区块: %s  长度=%d' % (sw, (sw[1] - sw[0]) if sw[0] is not None else 0))

if su[0] is None or sw[0] is None:
    print('!! 定位失败')
    sys.exit(1)

blk = U[su[0]:su[1]]
print()
print('搬运内容前 200 字符:')
print(repr(blk[:200]))
print('搬运内容 BONDS=%d FORMUL=%d' % (blk.count('BONDS'), blk.count('FORMUL')))

W2 = W[:sw[0]] + blk + W[sw[1]:]
open(OUT, 'w', encoding='utf-8', errors='ignore').write(W2)
print()
print('写出 %s  %d bytes' % (os.path.basename(OUT), os.path.getsize(OUT)))
V = open(OUT, encoding='utf-8', errors='ignore').read()
print('新文件 BONDS=%d FORMUL=%d MOLEC-STRUCT=%d' % (V.count('BONDS'), V.count('FORMUL'), V.count('MOLEC-STRUCT')))
i = V.find('MOLEC-STRUCT')
print('区块头部:', repr(V[i - 30:i + 200]))
