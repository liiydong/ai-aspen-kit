# -*- coding: utf-8 -*-
"""MCP stdio 冒烟测试：初始化 + 调用 product_describe"""
import json, subprocess, sys

PY = r'D:\<化工运行时>\python-env\Scripts\python.exe'
SRV = r'D:\<化工工作区>\chemical-engineering-runtime\tools\expert_mcp.py'

def send(p, obj):
    p.stdin.write(json.dumps(obj) + "\n")
    p.stdin.flush()

def recv(p, timeout_lines=400):
    for _ in range(timeout_lines):
        line = p.stdout.readline()
        if not line:
            return None
        line = line.strip()
        if line.startswith("{"):
            try:
                return json.loads(line)
            except ValueError:
                continue
    return None

p = subprocess.Popen([PY, '-X', 'utf8', SRV], stdin=subprocess.PIPE,
                     stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                     text=True, encoding='utf-8', bufsize=1)

send(p, {"jsonrpc": "2.0", "id": 1, "method": "initialize",
         "params": {"protocolVersion": "2024-11-05", "capabilities": {},
                    "clientInfo": {"name": "smoke", "version": "1.0"}}})
r = recv(p)
if r and 'result' in r:
    si = r['result'].get('serverInfo', {})
    print('OK  initialize ->', si.get('name'), si.get('version'))
else:
    print('FAIL initialize ->', str(r)[:300]); sys.exit(1)

send(p, {"jsonrpc": "2.0", "method": "notifications/initialized", "params": {}})
send(p, {"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}})
r = recv(p)
if r and 'result' in r:
    names = [t['name'] for t in r['result'].get('tools', [])]
    print('OK  tools/list ->', len(names), '个工具:', ', '.join(names))
else:
    print('FAIL tools/list ->', str(r)[:300])

send(p, {"jsonrpc": "2.0", "id": 3, "method": "tools/call",
         "params": {"name": "pressure_calculate",
                    "arguments": {"method": "series_pressure",
                                  "inputs": {"inlet_pressure_pa": 500000,
                                             "losses_pa": [20000, 30000]}}}})
r = recv(p)
if r and 'result' in r:
    txt = json.dumps(r['result'], ensure_ascii=False)
    print('OK  tools/call pressure_calculate ->', txt[:260])
else:
    print('FAIL tools/call ->', str(r)[:300])

p.stdin.close()
p.terminate()
print('冒烟测试结束')
