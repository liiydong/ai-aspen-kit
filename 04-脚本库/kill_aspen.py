# -*- coding: utf-8 -*-
"""通过 WMI 关闭所有 AspenPlus 进程"""
import time
import win32com.client as win32

wmi = win32.GetObject('winmgmts:\\\\.\\root\\cimv2')
procs = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='AspenPlus.exe'")
n = 0
for p in procs:
    try:
        print(f'终止 PID {p.ProcessId}  ({p.CommandLine[:80] if p.CommandLine else "无命令行"})')
        p.Terminate()
        n += 1
    except Exception as e:
        print('  失败:', str(e)[:120])
print(f'共终止 {n} 个进程')
time.sleep(3)

procs2 = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='AspenPlus.exe'")
left = sum(1 for _ in procs2)
print('剩余 AspenPlus 进程数:', left)
