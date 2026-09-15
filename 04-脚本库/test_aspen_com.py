# -*- coding: utf-8 -*-
"""测试 Aspen Plus V15 COM 接口连通性与许可"""
import time, sys
import win32com.client as win32

print('正在连接 Aspen Plus COM（首次会自动启动 Aspen，可能需 1~3 分钟）...')
sys.stdout.flush()
t = time.time()
try:
    app = win32.Dispatch('Apwn.Document')
    print('连接成功，用时 %.1f s' % (time.time() - t))
    for attr in ('Version', 'FullName', 'Name', 'Visible'):
        try:
            print(f'  {attr} =', getattr(app, attr))
        except Exception as e:
            print(f'  {attr} 读取失败:', e)
    # 探测常用接口
    try:
        tree = app.Tree
        print('  Tree 对象: OK')
    except Exception as e:
        print('  Tree 对象失败:', e)
    try:
        eng = app.Engine
        print('  Engine 对象: OK')
    except Exception as e:
        print('  Engine 对象失败:', e)
    print('RESULT: COM_OK')
except Exception as e:
    print('连接失败:', repr(e))
    print('RESULT: COM_FAIL')
