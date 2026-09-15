# -*- coding: utf-8 -*-
"""用 os.kill 终止 AspenPlus 进程"""
import os, signal, time
import win32com.client as win32

wmi = win32.GetObject('winmgmts:\\\\.\\root\\cimv2')
pids = [p.ProcessId for p in wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='AspenPlus.exe'")]
print('待终止 PID:', pids)
for pid in pids:
    try:
        os.kill(pid, signal.SIGTERM)
        print('  已发送终止信号:', pid)
    except Exception as e:
        print('  失败', pid, ':', str(e)[:120])
time.sleep(3)

left = [p.ProcessId for p in wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='AspenPlus.exe'")]
print('剩余 AspenPlus PID:', left if left else '无 ✓')
