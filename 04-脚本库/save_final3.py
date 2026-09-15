# -*- coding: utf-8 -*-
import re, shutil, os

SRC = r'D:\<化工工作区>\_probe\wv4_nphase.bkp'
DST = r'D:\<化工工作区>\NA-Chemical-10000t_定稿v3.bkp'
shutil.copy(SRC, DST)

t = open(DST, encoding='utf-8', errors='ignore').read()
print('size', len(t))
print('BLOCK 段:', len(re.findall(r'\?\s*BLOCK\s', t)))
print('RUN-CLASS:', re.findall(r'RUN-CLASS = (\w+)', t))
m = re.search(r'\?\s*SETUP\s+"SIM-OPTIONS"\s*\?.*?\?', t, re.S)
print('SIM-OPTIONS:', re.sub(r'\s+', ' ', m.group(0)) if m else '未找到')
print('组分条数:', len(re.findall(r'CID = ', t[t.find('COMPONENTS MAIN'):t.find('COMPONENTS "COMP-LIST"')])))
i = t.find('PARAMNAME = NRTL')
j = t.find('PARAMNAME =', i + 20)
seg = t[i:j] if j > i else t[i:i+40000]
print('UVAL =', seg.count('UVAL'), ' BPVAL =', seg.count('BPVAL'))
print('E-104 PRES:', re.findall(r'PRES = 1\.8', t)[:1], ' E-201 PRES:', re.findall(r'PRES = 1\.0 ', t)[:1])
# 流股数与表格完整性粗查
print('STREAM 段:', len(re.findall(r'\?\s*STREAM\s+MATERIAL', t)))
print('含 $_END:', '$_END_ADS_FILE' in t or '$' in t[-200:])
