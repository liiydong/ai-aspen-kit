# Aspen Plus 操作要点与踩坑总结

> 来源：烟酰胺项目从"模型打不开"到"36 单元 0 错误收敛"的全过程，以及二元交互参数攻坚。
> 目的：**同一类错误不再犯第二次**。

---

## 一、最容易搞混的一件事：Aspen 有两个"环境"

Aspen Plus 界面上有环境切换标签（`物性` / `模拟` / …）。**同一个 F5，在不同环境下做的完全不是一件事**：

| 环境 | F5 做什么 | 什么时候用 |
|---|---|---|
| **`物性`（Properties）** | 物性计算、**参数估算** | 补分子结构、估算二元参数 |
| **`模拟`（Simulation）** | 流程模拟 | 算流程、看结果 |

**踩过的坑**：让用户在 `模拟` 环境按 F5 做"估算"，结果卡在"运行又要求先有二元参数"的死循环 —— **不是参数的问题，是环境选错了**。
**记牢**：**估参数去 `物性`，跑流程去 `模拟`。**

---

## 二、文件格式：只用 `.bkp`

| 格式 | 本质 | 可文本编辑 | 备注 |
|---|---|---|---|
| `.bkp` | Aspen 输入语言（**纯文本**） | ✅ | **唯一可靠** |
| `.apwz` | ZIP 包，内含 `.apw`（二进制）+ `.bkp.backup` | ⚠️ 能解开 | **`.bkp.backup` 是降级备份**：缺 `DATABANKS`、`RUN-CLASS` 会退回 `PROP` |
| `.apw` | 二进制主文件 | ❌ | COM 打不开 `.apwz`，解出 `.bkp` 后才能开 |

**踩过的坑**：
1. 早期用 `.apwz` 交付 → 打开后**看不到流股/模块**（`RUN-CLASS = PROP`），白折腾好几轮。
2. **即使在 GUI 里"另存为 .bkp"，也要确认流程图还在** —— 用户在物性环境操作后另存的 `.bkp`，`\Data` 只有 **7** 个子节点，流程图又丢了。

**判据（记牢）**：健康的模型 `\Data` 应有 **25** 个子节点，且含 `Streams` / `Blocks` / `Flowsheet`。

---

## 三、`RUN-CLASS`：流程图加载的总开关

- 症状：COM 读 `\Data` 只有 8 个子节点，无 `Streams`/`Blocks`；流股参数报 `AE_UNDERSPEC`（规格不全）；模块无法实例化。
- 根因：`SETUP MAIN` 段里的 `RUN-CLASS = PROP`（仅物性）——**该模式根本不加载流程图**。
- 修法：改成 `RUN-CLASS = FLOWSHEET`。

---

## 四、树路径（这里浪费过最多时间）

| 用途 | 正确路径 | 易错点 |
|---|---|---|
| 二元交互参数 | `\Data\Properties\Parameters\Binary Interaction\NRTL-1` | **`Binary Interaction` 中间有空格**！写成 `...\Parameters\Binary\NRTL-1` 会返回 `None` |
| 分子结构 | `\Data\Properties\Molecular Structure\<组分>\Input` | 是 `Molecular Structure`（带空格），不是 `MOLEC-STRUCT` |
| 组分表 | `\Data\Components\Specifications\Input` | 其子节点是**列**（ANAME、CASN…），不是行 |
| 估算设置 | `\Data\Properties\Estimation\Estimate\Input` | 含 `ALLONLY`（估算全部缺失参数） |
| 数据库 | `\Data\Pure Databanks\Input\FILE_SYM_NAM`、`\Data\Other Databanks\Input\AUTO_PARAM` | — |

---

## 五、bkp 文本格式的四条铁律

1. **记录分隔符是单个 `\`**，且必须**独立占一行或以 `\ ` 起头**。
   - 把多条记录挤在一行 → Aspen 只解析第一条（表现：反应报"不守恒"、塔的 `FEEDS` 只进一条）。
2. **每行约 76 字符必须换行**。超长单行 → 整段解析失败（表现：模块报参数缺失）。
3. **改块类型必须同步改三处**：① 文件头的块实例注册表；② `FLOWSHEET GLOBAL` 段的 `BLKTYPE/MDLTYPE`；③ 块段落头。
4. **段头可能被折行拆开**，例如：
   ```
   ? PROPERTIES "MOLEC-STRUCT" \n"3-MP" ?
   ? PROPERTIES "MOLEC-STRUCT" CO \n?
   ```
   → **不要用"正则匹配组分名"定位段**，要**按 `MOLEC-STRUCT` 出现位置做整块切分**。
   （本次因此返工 3 次，最终用位置切分才成功。）

---

## 六、COM 能做与不能做

| 操作 | COM | 说明 |
|---|---|---|
| 读/写参数值 | ✅ | 改完**必须回读**确认 |
| 运行、取数、导出 | ✅ | — |
| 新增**流股** | ✅ | `\Data\Streams\Elements.Add("S-101")` |
| 新增**组分** | ❌ | `Elements.Add` → `Aspen.Navigation：该功能尚未生成` |
| 写**分子结构**（原子/键） | ❌ | 同上；`Tree.NewChild` 只接受 1 个位置参数 |
| 触发**数据库结构检索** | ❌ | 写 `CC Nodes\HASFORMUL/GEN/FORMULA` 报 `AE_UNDERSPEC`/`AE_UNKERR` |

**结论**：**加组分、写结构、触发检索这三件事必须走图形界面。**

---

## 七、二元交互参数（本项目最大的坑）

### 诊断顺序（按序查，别跳步）
1. `NRTL-1\Output\CID1` 的行数 = **实际可用参数条数**。
2. **警惕"空壳记录"**：bkp 里可能出现
   ```
   BPVAL PARAMNAME2 = NRTL CID1 = H2O CID2 = TOL ... VAL1 = "APV150 LLE-ASPEN" ...
   ```
   只有 `VAL`（来源标记）**没有 `UVAL`（数值）** → **对计算无贡献**，别被"有 1 条参数"误导。
   **而且它有害**：实测删掉它之后模型才能正常运行。
3. 缺参数的原因只有两种：**库里没有实测数据** / **需要 UNIFAC 估算**。
4. UNIFAC 估算的前置：`UNIFAC-Groups\Input\GROUPNO` 有行 ← 依赖**分子结构**。

### 分子结构（关键突破点）
- **结构数据本来就在 Aspen 数据库里**，只是文件里的表是空的。
- GUI 路径：`物性 → 分子结构 → 点组分 → 常规页（会显示分子图）→ 点 `键能计算`（英文 `Calculate Bonds`）`
- 点完结构会写进 bkp，格式：
  ```
  ? PROPERTIES "MOLEC-STRUCT" TOL ?
  \ BONDS ATOM1 = 1 DISPATOM1 = C ATOM2 = 2 DISPATOM2 = C BONDTYPE = D / ...
  \ FORMUL ATOMNO = C NOATOM = 7. <0> <0> / ...
  ```
  键级：`S`=单键、`D`=双键、`T`=三键。

### 即使补了结构，UNIFAC 仍可能失败
实测报错：
```
FUNCTIONAL GROUP GENERATION FOR THE UNIFAC METHOD CANNOT BE COMPLETED
FOR COMPONENT NAM.  THE FOLLOWING ATOMS WERE NOT MATCHED:
C 7  N 8  O 9
```
→ **烟酰胺（NAM）的酰胺基无法被 Aspen 的 UNIFAC 基团表识别**。
**下一步方向**：换 `UNIF-DMD`（Dortmund 修正版，基团表更全）而不是默认 `UNIFAC`。

---

## 八、单位与读数

- **判断某个输出量的量纲，最可靠的办法是切单位集重跑**：`\Data\Setup\Global\Input\INSET` 设为 `'SI'`，重跑后读同一节点即可换算。
  实测：`METCBAR` 下 `COND_DUTY/REB_DUTY` 读数为 **Gcal/h**（1 Gcal/h = 1163 kW）。
- 常用节点：流股 `Output\MASSFLMX\MIXED`、`MOLEFLMX\MIXED`、`TEMP_OUT\MIXED`、`PRES_OUT\MIXED`、`MASSFLOW3`；
  塔 `TOP_TEMP`、`BOTTOM_TEMP`、`COND_DUTY`、`REB_DUTY`。

---

## 九、`.bkp` 不保存计算结果

必须"**打开 → 运行 → 同一进程内读取/导出/另存**"，否则下次打开结果全空。

---

## 十、捕获报错文本

`Engine.ControlPanel` 在本机不可用 → **改用 COM 事件钩子**：

```python
import win32com.client as win32, pythoncom
MSGS = []
class Sink:
    def OnControlPanelMessage(self, *a):
        s = ' '.join(str(x) for x in a).strip()
        if s and s != 'False':
            MSGS.append(s)

doc = win32.DispatchEx('Apwn.Document')
win32.WithEvents(doc, Sink)
doc.InitFromArchive2(bkp_path)
doc.Engine.Run2(False)          # 异步！
while doc.Engine.IsRunning:
    pythoncom.PumpWaitingMessages(); time.sleep(0.25)
for _ in range(80):             # 收尾消息
    pythoncom.PumpWaitingMessages(); time.sleep(0.05)
```
**要点**：`Run2` 是异步的，必须 `PumpWaitingMessages()` + 轮询 `IsRunning`，否则拿不到结果。

---

## 十一、流程搭建与标定经验

- **加设备**：bkp 文本注入 + 同步 FLOWSHEET + 块段落，三处缺一不可；每加一批就重跑，用 `Terminal/Severe/Errors` 计数当判据。
- **反应器（RStoic）**：反应可内嵌在模块里（`STOIC`/`CONVEX`），**反应物用 `STOIC`、产物必须用 `STOIC1`**；配平错误 Aspen 会给出差值（如 `-12.011` → 差 1 个碳）。
- **端口语义**：`Flash2` = `M0-1=V(气相)`、`M1-2=L(液相)`；`RadFrac` = `M1-2=塔顶`、`M2-3=塔底`；`Sep` 全部 `M0-1`。
- **D:F（馏出/进料）**：清晰分割时应等于**塔顶组分的摩尔分率**，设错会把产品打进塔釜。
- **降回流比**通常不改变分离结果（只要纯度仍达标），但**显著降低再沸负荷**（实测降 42%）。
- **缺二元参数的分离单元**（萃取、同分异构体），用 `Sep` 按实测分离率等效建模，并在文档中如实说明。

---

## 十二、交付前自检清单

- [ ] `RUN-CLASS = FLOWSHEET`（`\Data` 有 25 个子节点）
- [ ] 文件名用 **ASCII**（中文名在 `SaveAs` 时偶发失败）
- [ ] 至少跑一次，`Terminal / Severe / Errors` 全为 0
- [ ] 总物料平衡**进=出**，偏差 < 0.01%
- [ ] 参数改动**回读确认**（不能只看文本写成功）
- [ ] 结果在**同一进程内**导出（`.bkp` 不存结果）
- [ ] 交付 `.apwz`（给人双击看）+ `.bkp`（给程序读）双份
