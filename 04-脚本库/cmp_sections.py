# -*- coding: utf-8 -*-
"""对比段结构：找出我们缺的段"""
import re

files = {
    '参考(能跑)': r'D:\<化工工作区>\_probe\ref_multi.bkp',
    '我们': r'D:\<化工工作区>\_probe\noair2.bkp',
}
secs = {}
for tag, p in files.items():
    t = open(p, encoding='utf-8', errors='ignore').read()
    lst = []
    for m in re.finditer(r'\?\s*([^?\\]{1,45}?)\s*\?', t):
        s = re.sub(r'\s+', ' ', m.group(1).strip())
        if s:
            lst.append(s)
    secs[tag] = lst
    print('=== %s：段数 %d ===' % (tag, len(lst)), flush=True)
    print('   ', lst[:60], flush=True)
    print(flush=True)

r = set(secs['参考(能跑)'])
u = set(secs['我们'])
print('=== 参考有、我们没有的段 ===', flush=True)
for x in sorted(r - u):
    print('   ', x, flush=True)
print()
print('=== 我们有、参考没有的段 ===', flush=True)
for x in sorted(u - r):
    print('   ', x, flush=True)

# 参考的 STREAM-CLASS 相关段内容
t = open(files['参考(能跑)'], encoding='utf-8', errors='ignore').read()
for pat in [r'\?\s*"STREAM-CLASS"\s+SUBSTREAMS\s*\?',
            r'\?\s*"STREAM-CLASS"\s+"STREAM-CLASS"\s*\?',
            r'\?\s*"STREAM-CLASS"\s+"DEF-STREAMS"\s*\?']:
    m = re.search(pat, t)
    if not m:
        print('未找到', pat, flush=True)
        continue
    m2 = re.search(r'\?\s+\S', t[m.end():])
    end = m.end() + (m2.start() if m2 else 200)
    print()
    print('=== 参考的', re.sub(r'\s+', ' ', m.group()), '===', flush=True)
    print(re.sub(r'\s+', ' ', t[m.start():end])[:700], flush=True)
