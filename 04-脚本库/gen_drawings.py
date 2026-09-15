# -*- coding: utf-8 -*-
"""生成附录 C 的 BFD 与 PFD（SVG 矢量图）"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
OUTDIR = r'D:\obsidian-vault\毕业设计\化学法版'

C_BLUE = '#1f4e79'
C_BAND = '#eaf1f8'
C_BOX = '#ffffff'
C_REACT = '#fdece7'
C_SEP = '#e8f3ec'
C_TXT = '#1a1a1a'


def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


# ============================ BFD ============================
def bfd():
    L = []
    W, H = 1500, 900
    L.append('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 %d %d" font-family="SimHei,Microsoft YaHei,sans-serif">' % (W, H))
    L.append('<rect width="%d" height="%d" fill="#fbfcfe"/>' % (W, H))
    L.append('<text x="750" y="46" font-size="26" font-weight="bold" fill="%s" text-anchor="middle">年产 10000 吨烟酰胺装置 总方块流程图（BFD）</text>' % C_BLUE)
    L.append('<text x="750" y="72" font-size="14" fill="#555" text-anchor="middle">化学法（氨氧化 → 吸收 → 萃取 → 精馏 → 化学水解 → 精制）　共 36 个单元　图号 C-1</text>')

    # 工段方块
    stages = [
        ('3-甲基吡啶\n液氨、空气', 60, 150, '#f2f2f2'),
        ('① 氨氧化\nR-101（405 ℃）', 250, 150, C_REACT),
        ('② 急冷吸收\nT-201 / T-202', 440, 150, C_SEP),
        ('③ 甲苯萃取\nT-301 / T-302', 630, 150, C_SEP),
        ('④ 四塔精馏\nT-401~T-404', 820, 150, C_SEP),
        ('⑤ 化学水解\nR-501（140 ℃）', 1010, 150, C_REACT),
        ('⑥ 精制\n蒸发/结晶/干燥', 1200, 150, C_SEP),
    ]
    for txt, x, y, col in stages:
        L.append('<rect x="%d" y="%d" width="160" height="86" rx="8" fill="%s" stroke="%s" stroke-width="1.6"/>' % (x, y, col, C_BLUE))
        for i, line in enumerate(txt.split('\n')):
            L.append('<text x="%d" y="%d" font-size="14" fill="%s" text-anchor="middle">%s</text>'
                     % (x + 80, y + 34 + i * 20, C_TXT, esc(line)))
    # 主流程箭头
    for i in range(len(stages) - 1):
        x0 = stages[i][1] + 160
        L.append('<line x1="%d" y1="193" x2="%d" y2="193" stroke="%s" stroke-width="2.2" marker-end="url(#ar)"/>'
                 % (x0, stages[i + 1][1] - 4, C_BLUE))
    # 主产品
    L.append('<rect x="1200" y="300" width="160" height="60" rx="8" fill="#fff5e6" stroke="#c47f00" stroke-width="1.6"/>')
    L.append('<text x="1280" y="326" font-size="14" fill="#7a4a00" text-anchor="middle">烟酰胺产品</text>')
    L.append('<text x="1280" y="346" font-size="13" fill="#7a4a00" text-anchor="middle">1419.5 kg/h（99.62%）</text>')
    L.append('<line x1="1280" y1="236" x2="1280" y2="296" stroke="#c47f00" stroke-width="2.2" marker-end="url(#ar2)"/>')

    # 副产与排出
    outs = [('尾气 S-130\n5805.9 kg/h', 440, 300, '#dbe7f3'), ('废水 S-213\n2037.3 kg/h', 630, 300, '#dbe7f3'),
            ('副产 4-CP S-119\n6.6 kg/h', 820, 300, '#f3e6f7'), ('三效冷凝水\n2695.2 kg/h', 1010, 300, '#dbe7f3'),
            ('母液排放 S-405C\n93.9 kg/h', 1010, 390, '#dbe7f3'), ('回收甲苯 S-115\n2346.2 kg/h', 630, 390, '#e6f3e6')]
    for txt, x, y, col in outs:
        L.append('<rect x="%d" y="%d" width="160" height="60" rx="6" fill="%s" stroke="#8aa4bd" stroke-width="1.2"/>' % (x, y, col))
        for i, line in enumerate(txt.split('\n')):
            L.append('<text x="%d" y="%d" font-size="12" fill="#33475b" text-anchor="middle">%s</text>'
                     % (x + 80, y + 26 + i * 17, esc(line)))

    # 循环回路
    L.append('<path d="M 950 360 L 950 470 L 300 470 L 300 240" fill="none" stroke="#0e7a3c" stroke-width="2" stroke-dasharray="7 5" marker-end="url(#ar3)"/>')
    L.append('<text x="640" y="490" font-size="13" fill="#0e7a3c" text-anchor="middle">循环：未反应 3-MP（S-117M，14.5 kg/h）返回反应器</text>')
    L.append('<path d="M 1290 400 L 1290 540 L 700 540 L 700 250" fill="none" stroke="#0e7a3c" stroke-width="2" stroke-dasharray="7 5" marker-end="url(#ar3)"/>')
    L.append('<text x="990" y="560" font-size="13" fill="#0e7a3c" text-anchor="middle">循环：结晶母液（S-405B，1079.8 kg/h，92%）返回脱色/蒸发工段</text>')
    L.append('<path d="M 800 236 L 800 660 L 900 660 L 900 250" fill="none" stroke="#0e7a3c" stroke-width="2" stroke-dasharray="7 5" marker-end="url(#ar3)"/>')
    L.append('<text x="1010" y="680" font-size="13" fill="#0e7a3c" text-anchor="middle">循环：甲苯（S-117T，80.1 kg/h）返回萃取塔</text>')

    # 物料平衡框
    L.append('<rect x="60" y="720" width="1380" height="120" rx="8" fill="#f7f9fc" stroke="#8aa4bd"/>')
    L.append('<text x="80" y="748" font-size="15" font-weight="bold" fill="%s">总物料平衡（Aspen Plus V15 全流程模拟校核）</text>' % C_BLUE)
    L.append('<text x="80" y="776" font-size="13" fill="#333">进料 15006.04 kg/h ＝ 出料 15006.04 kg/h，偏差 0.0000%；Terminal / Severe / Errors 均为 0</text>')
    L.append('<text x="80" y="800" font-size="13" fill="#333">产品 1419.5 kg/h（烟酰胺 99.62 wt%、水分 0.28 wt%）＝ 10221 t/a（设计 10000 t/a 的 102.2%）</text>')
    L.append('<text x="80" y="824" font-size="13" fill="#333">3-甲基吡啶单耗 910.7 kg/t、总收率 83.73%（设计 936.89 kg/t、82.25%）</text>')

    L.append('<defs><marker id="ar" markerWidth="9" markerHeight="9" refX="7" refY="3" orient="auto"><path d="M0,0 L7,3 L0,6 z" fill="%s"/></marker>' % C_BLUE)
    L.append('<marker id="ar2" markerWidth="9" markerHeight="9" refX="7" refY="3" orient="auto"><path d="M0,0 L7,3 L0,6 z" fill="#c47f00"/></marker>')
    L.append('<marker id="ar3" markerWidth="9" markerHeight="9" refX="7" refY="3" orient="auto"><path d="M0,0 L7,3 L0,6 z" fill="#0e7a3c"/></marker></defs>')
    L.append('</svg>')
    open(os.path.join(OUTDIR, 'C-1_BFD.svg'), 'w', encoding='utf-8').write('\n'.join(L))
    print('C-1_BFD.svg 已生成')


# ============================ PFD ============================
def pfd():
    ROWS = [
        ('反应与吸收工段', C_REACT, [
            ('C-101', '空气鼓风机', 'S-103 5401.1 → S-127'),
            ('E-102', '空气预热器', 'S-127 → S-128 150℃'),
            ('E-103', '氨汽化器', 'S-102 243.9 → S-126'),
            ('M-101', '原料混合器', '+S-117M 14.5 循环'),
            ('E-101', '原料气预热器', 'S-104 → S-105'),
            ('R-101', '氨氧化反应器', '405℃  S-106'),
            ('E-105', '废热锅炉', '回收蒸汽 11.8 t/h'),
            ('E-104', '反应气冷却器', 'S-129 → S-108'),
        ]),
        ('急冷吸收与萃取工段', C_SEP, [
            ('T-201', '急冷吸收塔', 'S-107 2334.3 水'),
            ('T-202', '尾气洗涤塔', 'S-130 5805.9 排空'),
            ('E-201', '吸收液冷却器', 'S-110 → S-112 40℃'),
            ('T-301', '甲苯萃取塔', 'S-111 2322.3 甲苯'),
            ('P-301', '萃取相泵', 'S-114 → S-114A'),
            ('T-302', '萃余水汽提塔', 'S-213 2037.3 废水'),
        ]),
        ('四塔精馏工段', C_SEP, [
            ('T-401', '脱甲苯塔', '顶 53.2℃/0.24 bar'),
            ('P-401', '粗 3-CP 泵', 'S-116 → S-116A'),
            ('T-402', '脱轻组分塔', 'D:F = 0.070'),
            ('M-504', '甲苯回收器', '3-MP/甲苯分流'),
            ('T-403', '脱 4-CP 塔', 'S-119 6.6 副产'),
            ('T-404', '3-CP 产品塔', 'D:F = 0.990  S-121'),
        ]),
        ('化学水解工段', C_REACT, [
            ('R-501', '化学水解釜', '140℃  转化率 99%'),
            ('E-501', '水解液冷却器', '140 → 80℃'),
        ]),
        ('精制工段（脱色—蒸发—结晶—离心—干燥）', C_SEP, [
            ('M-501', '脱色釜', '+S-405B 1079.8 循环'),
            ('M-502', '板框压滤机', 'S-211 29.6 废炭'),
            ('E-601', '一效蒸发器', '1.0 bar  蒸发 1111.9'),
            ('E-604', '一效冷凝器', 'S-216'),
            ('E-602', '二效蒸发器', '0.47 bar  蒸发 908.3'),
            ('E-605', '二效冷凝器', 'S-217'),
            ('E-603', '三效蒸发器', '0.20 bar  蒸发 675.0'),
            ('E-606', '三效冷凝器', 'S-218'),
            ('C-601', '结晶釜', '70.89 → 25℃  收率 75%'),
            ('M-601', '离心机', 'S-208 湿品 1492.5'),
            ('P-501', '母液泵', 'S-405 1173.7'),
            ('M-503', '母液分流器', '92% 循环 / 8% 排放'),
            ('D-601', '气流干燥器', '130℃/0.05 bar  S-209 产品'),
            ('P-302', '废水泵', 'S-213A 送处理'),
        ]),
    ]
    W = 1720
    BW, GAP = 152, 12
    ROWINFO = []
    for title, band, blocks in ROWS:
        n = len(blocks)
        if n * BW + (n - 1) * GAP > 1660:
            per = n // 2 + n % 2
            sub = [blocks[:per], blocks[per:]]
        else:
            sub = [blocks]
        ROWINFO.append((title, band, sub, 128 if len(sub) == 1 else 190))
    H = 96 + sum(r[3] for r in ROWINFO) + 90
    L = []
    L.append('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 %d %d" font-family="SimHei,Microsoft YaHei,sans-serif">' % (W, H))
    L.append('<rect width="%d" height="%d" fill="#fbfcfe"/>' % (W, H))
    L.append('<text x="%d" y="40" font-size="25" font-weight="bold" fill="%s" text-anchor="middle">年产 10000 吨烟酰胺装置 工艺物料流程图（PFD）</text>' % (W // 2, C_BLUE))
    L.append('<text x="%d" y="66" font-size="13.5" fill="#555" text-anchor="middle">化学法　36 个单元 / 61 条流股　物料平衡偏差 0.0000%%　图号 C-2　（流量单位 kg/h）</text>' % (W // 2))

    y = 96
    for title, band, rows, row_h in ROWINFO:
        n = sum(len(r) for r in rows)
        L.append('<rect x="20" y="%d" width="%d" height="%d" rx="10" fill="%s" stroke="#c9d8e6"/>' % (20, y, row_h - 14, band))
        L.append('<text x="34" y="%d" font-size="15" font-weight="bold" fill="%s">%s</text>' % (y + 24, C_BLUE, esc(title)))
        x0 = 34
        yy = y + 40
        bw, gap = BW, GAP
        for r in rows:
            x = x0
            for (bid, name, note) in r:
                L.append('<rect x="%d" y="%d" width="%d" height="52" rx="6" fill="%s" stroke="%s" stroke-width="1.4"/>'
                         % (x, yy, bw, C_BOX, C_BLUE))
                L.append('<text x="%d" y="%d" font-size="12.5" font-weight="bold" fill="%s" text-anchor="middle">%s</text>'
                         % (x + bw // 2, yy + 20, C_BLUE, esc(bid)))
                L.append('<text x="%d" y="%d" font-size="11.5" fill="%s" text-anchor="middle">%s</text>'
                         % (x + bw // 2, yy + 37, C_TXT, esc(name)))
                if note:
                    L.append('<text x="%d" y="%d" font-size="10" fill="#666" text-anchor="middle">%s</text>'
                             % (x + bw // 2, yy + 48, esc(note)))
                if x + bw + gap < x0 + 1660:
                    L.append('<line x1="%d" y1="%d" x2="%d" y2="%d" stroke="%s" stroke-width="1.8" marker-end="url(#a1)"/>'
                             % (x + bw, yy + 26, x + bw + gap - 2, yy + 26, C_BLUE))
                x += bw + gap
            yy += 62
        y += row_h

    L.append('<defs><marker id="a1" markerWidth="7" markerHeight="7" refX="5.5" refY="2.5" orient="auto">'
             '<path d="M0,0 L5.5,2.5 L0,5 z" fill="%s"/></marker></defs>' % C_BLUE)
    # 图签
    ty = H - 74
    L.append('<rect x="1140" y="%d" width="560" height="60" fill="none" stroke="#8aa4bd"/>' % ty)
    L.append('<line x1="1140" y1="%d" x2="1700" y2="%d" stroke="#8aa4bd"/>' % (ty + 30, ty + 30))
    L.append('<line x1="1420" y1="%d" x2="1420" y2="%d" stroke="#8aa4bd"/>' % (ty, ty + 60))
    L.append('<text x="1152" y="%d" font-size="12" fill="#333">设计项目：年产 10000 t 烟酰胺（化学法）</text>' % (ty + 20))
    L.append('<text x="1152" y="%d" font-size="12" fill="#333">校核：Aspen Plus V15 全流程模拟</text>' % (ty + 50))
    L.append('<text x="1432" y="%d" font-size="12" fill="#333">图号：C-2</text>' % (ty + 20))
    L.append('<text x="1432" y="%d" font-size="12" fill="#333">比例：示意</text>' % (ty + 50))
    L.append('</svg>')
    open(os.path.join(OUTDIR, 'C-2_PFD.svg'), 'w', encoding='utf-8').write('\n'.join(L))
    print('C-2_PFD.svg 已生成')


bfd()
pfd()
print('DONE')
