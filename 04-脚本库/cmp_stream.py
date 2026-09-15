import re, sys
sys.stdout.reconfigure(encoding='utf-8')

def stream_para(tag, path, sids):
    t = open(path, encoding='utf-8', errors='ignore').read()
    print('=' * 30, tag)
    for sid in sids:
        pat = re.compile(r'\?\s*STREAM\s+MATERIAL\s+"?%s"?\s*\?' % re.escape(sid))
        ms = list(pat.finditer(t))
        print('  --- %s  段落数=%d' % (sid, len(ms)))
        for m in ms:
            nxt = re.search(r'\n\?\s', t[m.end():])
            seg = t[m.start():m.end() + (nxt.start() if nxt else 400)]
            print('     ', re.sub(r'\s+', ' ', seg)[:420])
    print()

stream_para('原生 多效精馏', r'D:/<化工工作区>/_probe/ref_multi.bkp', ['0401', 'S4', 'S8'])
stream_para('我们的 sim6', r'D:/<化工工作区>/NA-Chemical-10000t_sim6.bkp',
            ['S-101', 'S-107', 'S-111'])
