# -*- coding: utf-8 -*-
"""穷举让 Appen 真正跑起来的方式"""
import time
import win32com.client as win32

F = r'D:\<化工工作区>\NA-Chemical-10000t_模拟.bkp'


def check(doc, tag, wait=25):
    t0 = time.time()
    while time.time() - t0 < wait:
        time.sleep(2.5)
        try:
            if doc.Engine.IsRunning is False:
                break
        except Exception:
            break
    out = []
    for s in ['S-104', 'S-110', 'S-201']:
        n = doc.Tree.FindNode(r'\Data\Streams\%s\Output\TEMP' % s)
        out.append('%s=%s' % (s, n.Value if n is not None else None))
    try:
        ready = doc.Engine.Ready
    except Exception:
        ready = '?'
    print('   %-28s Ready=%-6r %s' % (tag, ready, ' '.join(out)), flush=True)


CASES = []


def case(tag):
    def deco(fn):
        CASES.append((tag, fn))
        return fn
    return deco


@case('doc.Run2()')
def _(doc):
    doc.Run2()


@case('doc.Run()')
def _(doc):
    doc.Run()


@case('Engine.Reinit + Engine.Run2')
def _(doc):
    doc.Engine.Reinit()
    time.sleep(3)
    doc.Engine.Run2()


@case('Visible=True + Engine.Run2')
def _(doc):
    doc.Visible = True
    time.sleep(5)
    doc.Engine.Run2()


for tag, fn in CASES:
    print('=' * 18, tag, flush=True)
    doc = win32.DispatchEx('Apwn.Document')
    try:
        doc.SuppressDialogs = True
    except Exception:
        pass
    try:
        doc.InitFromArchive(F)
    except Exception as e:
        print('  打开失败', str(e)[:80], flush=True)
        continue
    time.sleep(4)
    try:
        fn(doc)
        print('   调用成功', flush=True)
    except Exception as e:
        print('   调用异常:', str(e)[:110], flush=True)
    check(doc, tag)
    try:
        doc.Close()
    except Exception:
        pass
    print(flush=True)
