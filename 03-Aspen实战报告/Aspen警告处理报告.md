# Aspen 警告处理报告

> 日期：2026-09-15
> 依据：你导出的计算报告 `ASPEN PLUS -2026.xlsx`（10:16 运行）＋ 我按当前模型重跑的控制面板原文（10:41 运行）
> 处置后模型：`NA-Chemical-10000t_定稿v3.bkp`

---

## 一、先核对你在报告里发现的两点

### 1. E-104、E-201 的压力 —— 你补的值已生效

| 设备 | 规格 | 指定温度 | 指定压力 | 出口温度 | 出口压力 | 热负荷 |
|---|---|---|---|---|---|---|
| **E-104** | 两相 TP 闪蒸 | **100.0 ℃** | **1.8 bar** | 100.00 ℃ | 1.8000 bar | −0.57701 Gcal/h |
| **E-201** | 两相 TP 闪蒸 | **40.0 ℃** | **1.0 bar** | 40.00 ℃ | 1.0000 bar | −0.17110 Gcal/h |

核对无误：① 出口流股 S-108 实测 100.0000 ℃/1.8000 bar、S-112 实测 40.0000 ℃/1.0000 bar，与各自 RESULTS 完全一致；② **补入的压力与上游一致**（S-129 进口 1.8 bar → E-104 出口 1.8 bar；S-110 进口 1.0 bar → E-201 出口 1.0 bar），无压降，物理上自洽。这两处改动正确。

顺带把全部 10 台 Heater 的压力/温度也核了一遍，**没有第二台缺压力**：E-101(1.8)、E-102(2.0)、E-103(2.0)、E-104(1.8)、E-105(1.8)、E-201(1.0)、E-501(1.0)、E-604(1.0)、E-605(0.47)、E-606(0.20)。

> 报告里没有 E-202 —— 你记的"另外那个"应该就是 T-202，它是**精馏塔（RADFRAC）**，不是换热器，不需要压力输入项。

### 2. 报告说"有软件完成但带警告"的模块 = **T-301、T-202**

报告第 96 页原文：

```
* Calculations were completed with warnings
* The following Unit Operation blocks were completed with warnings:
* T-301     T-202
* All streams were flashed normally
* All Convergence blocks were completed normally
```

我重跑后得到的结论完全一致，且把 20 条警告的**逐条原文**都抓出来了（COM 事件钩子）。

---

## 二、20 条警告的完整清单（按类别）

| 类别 | 条数 | 原文 | 归属 |
|---|---|---|---|
| **A. 首次迭代零流量旁路** | 5 | `MIXED SUBSTREAM HAS ZERO FLOW. BLOCK BYPASSED.` / `ZERO FEED - BLOCK BYPASSED.` / `ZERO FEED TO THE BLOCK. BLOCK BYPASSED` | P-301、T-401、P-401、T-402、M-504、M-101 |
| **B. T-301 液相稳定性** | 14 | `STABILITY CHECK FOR LIQUID PHASE: STABILITY CRITERION NOT SATISFIED FOR COMPONENT "H2O". STABILITY INDEX IS (1.1548), MOLE FRACTION IS (0.7558). THIS PHASE MAY BE UNSTABLE. TRY USING 3-PHASE (OR FREE-WATER, IF APPLICABLE) CALCULATIONS.` | T-301（SEP） |
| **C. T-202 中 O₂ 不在相平衡** | 1 | `WARNING WHILE CHECKING PHASE EQUILIBRIUM RESULTS FOR UNIT OPERATIONS / BLOCK: "T-202" (MODEL: "RADFRAC") / COMPONENT "O2" IN STREAM "S-130" IS NOT IN PHASE EQUILIBRIUM. RELATIVE ERROR IN VL FUGACITY IS 0.2969. FUGACITIES ARE 5131.0324 FOR VAPOR AND 6654.6033 FOR LIQUID. MOLE FRACTIONS ARE 5.1310-02 FOR VAPOR AND 1.0806D-10 FOR LIQUID. FUGACITY COEFFICIENTS ARE 1.0000 FOR VAPOR AND 6.1580D+08 FOR LIQUID.` | T-202（RADFRAC） |

合计 20 条 = 5 + 14 + 1，与 Aspen 汇总行的 `Warnings 0 0 20` 完全对上。

---

## 三、根因分析与处置

### 类别 A：首次迭代的零流量旁路（5 条）—— 无害，且**无法也无需消除**

这 5 条全部出现在 `$OLVER01` 的**第一次**迭代里（紧跟 `Converging tear streams: S-114` 之后）。原因是循环回路的撕裂流 S-114 在第一次代入时还没有数值，因此 P-301→T-401→P-401→T-402→M-504 这一串模块拿到的是零流量，Aspen 按"旁路"处理。

**第 2 次迭代起不再出现**（我逐条核对了 14 次迭代的消息，只在第 1 次出现）。这是**序贯模块法 + 循环回路模型的固有初始化提示**，任何带循环的 Aspen 模型都会打印，不代表模型有缺陷。

### 类别 B：T-301 液相稳定性（14 条）—— **已解决**

**根因**：T-301 是甲苯萃取塔，它的**混合进料本身就是一个两液相体系**——水相（水 + 溶解的 3-CP）与有机相（甲苯 + 3-CP）。我把三个进料加起来算了一下：

| 进料 | 水 kmol/h | 甲苯 kmol/h | 3-CP kmol/h | 合计 | **水摩尔分率** |
|---|---|---|---|---|---|
| S-112（水相）+ S-111（甲苯）+ S-117T | ≈127.6 | ≈25.2 | ≈12.2 | ≈165 | **0.773** |

与 Aspen 报的 `MOLE FRACTION IS (0.7558)` 基本吻合 —— 也就是说，**被标记的正是萃取塔的混合进料**。而 SEP 模块对进料做的是**单液相闪蒸**，所以稳定性检验判定"该液相可能不稳定"。Aspen 在提示里也直接给出了正确做法：`TRY USING 3-PHASE (OR FREE-WATER) CALCULATIONS`。

**处置**：在 `SETUP "SIM-OPTIONS"` 里开启全局三相闪蒸：

```
? SETUP "SIM-OPTIONS" ? ; "METCBAR_MOLE" ; \ "SIM-OPTIONS" NPHASE = 3 NPHASE-HI = 3 NPHASE-RB = 3 \ \ ?
```

**效果（已实测）**：**警告 20 → 6**，14 条 T-301 稳定性警告全部消除；同时

| 指标 | 处置前 | 处置后 |
|---|---|---|
| Terminal / Severe / Errors | 0 / 0 / 0 | **0 / 0 / 0** |
| 总物料平衡偏差 | 0.0000% | **0.0000%** |
| 产品 S-209 | 1419.249 kg/h | **1419.249 kg/h（不变）** |
| 萃余水 S-113 | 2173.9 kg/h | **2173.9 kg/h（不变）** |

即：**这是一个"把物理事实告诉 Aspen"的修正，不改变任何结果**。

### 类别 C：T-202 中 O₂ 不在相平衡（1 条）—— 需在界面做一步

**根因**：T-202 是尾气水洗塔，塔顶气 S-130 含 O₂ 5.13%（摩尔）。Aspen 报的液相逸度系数是 **6.1580×10⁸** —— 这个数在物理上不可能。原因是：**O₂ 是接近/超过临界温度的低沸气体，而 NRTL 是活度系数模型，本就不该用来算它的液相溶解**。正解是把这类难溶气体设为**亨利组分**，让 Aspen 用亨利定律而非活度系数处理它们的溶解。

**这一步我试过用脚本做，但没成功**：COM 不允许新增组分、也不接受我把 `HENRY-COMPS` 段按示例格式注入（两种段名写法都被 Aspen 忽略或拒绝）。这与我之前遇到的情况一致 —— **组分表的任何改动都必须经图形界面**。

**界面操作（约 1 分钟）**：

1. 打开模型 → 左侧 `组分` → **`亨利组分`（Henry-Comps）**
2. 在 `组分` 列依次选入（或输入）：**N2、O2、CO、CO2、HCN**
3. 按 F5 运行
4. `文件` → `另存为` → 类型选 **`Aspen Plus Backup (*.bkp)`** → 存回给我

预期：这 1 条警告消失，警告总数降到 **5**（即类别 A 那些无法消除的初始化提示）。

> 若不做这一步也完全可以：这条警告是"检查结果"的提示，说明 O₂ 的液相溶解量本该由亨利定律描述；模型里 O₂ 在液相的量是 1.08×10⁻¹⁰（摩尔分率），**对物料与热量衡算没有任何影响**。论文中已按"溶解气体按亨利定律处理的简化"如实说明即可。

---

## 四、处置结果总览

| 指标 | 你导出的报告 | 我的处置后模型 |
|---|---|---|
| Terminal Errors | 0 | **0** |
| Severe Errors | 0 | **0** |
| Errors | 0 | **0** |
| **Warnings** | **20** | **6** |
| 带警告的模块 | T-301、T-202 | 仅 T-202 |
| 总物料平衡 | 15006.04 = 15006.04（0.0000%） | **同上，不变** |
| 产品 S-209 | 1419.2487 kg/h | **1419.249 kg/h** |
| 外部进料 / 产品 | 7 条 / 15 条 | 不变 |
| 收敛块 $OLVER01 / $OLVER02 | 均 CONVERGED | 均 CONVERGED |

---

## 五、交付

| 文件 | 说明 |
|---|---|
| `NA-Chemical-10000t_定稿v3.bkp` | 修正后模型（36 单元 / 61 流股 / 15 组分；含你补的 E-104、E-201 压力；含 5 对真实二元参数；含 NPHASE=3） |
| 桌面 `烟酰胺Aspen模型_定稿v3_警告已处理.bkp` | 同上，便于直接打开 |

**保留未动的**：你原来那份 `NA-Chemical-10000t_BIP2.bkp` 仍在 `D:\<化工工作区>\` 下，没有覆盖。
