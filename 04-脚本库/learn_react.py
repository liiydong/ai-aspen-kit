# -*- coding: utf-8 -*-
import re

p = r'C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples\Bulk Chemicals\pfdtut.bkp'
t = open(p, encoding='utf-8', errors='ignore').read()
print('文件大小', len(t))
print()

print('=== 段标记 ===')
for m in re.finditer(r'\? [^?\\]{1,40} \?', t):
    print('   %7d  %s' % (m.start(), re.sub(r'\s+', ' ', m.group().strip())))
print()

i = t.find('? REACTIONS')
j = t.find('? ', i + 12)
# 找下一个段起始
for m in re.finditer(r'\? [^?\\]{1,40} \?', t):
    if m.start() > i + 5:
        j = m.start()
        break
print('=== REACTIONS 段 (位置 %d-%d) ===' % (i, j))
print(t[i:j])
