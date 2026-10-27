# VBA 学习项目（面向 ATE 自动化）

一个用于**系统学习 VBA 语言**的代码仓库。目标不是背语法，而是把 VBA 学到能独立写 ATE（自动化测试设备）侧的脚本：读穴表、判 pass/fail、批量刷数据进 Excel 报告。

每章一个目录，每节一个 `.vb` 文件，文件之间靠注释互相引用。代码里的示例尽量贴 ATE 场景（引脚名、电压值、测试结果），避免写成 `a = 1 + 1` 这种和业务无关的玩具代码。

---

## 1. 快速开始

本项目**不是可直接运行的程序**，只有纯文本 `.vb` 文件——VBA 必须活在 Office 宿主里。

1. 打开 Excel，新建一个空白工作簿，另存为 `.xlsm`（启用宏的工作簿）
2. `Alt + F11` 打开 VBA 编辑器
3. 菜单 `插入 → 模块`，新建一个标准模块
4. 打开对应 `.vb` 文件，**全选、复制、粘贴**进模块
5. 把模块名改成和文件对应的名字（`1.4.数组` → `1_4_数组`，**VBA 模块名不能带点**）
6. 光标放进任意 `Sub` 里按 `F5` 单跑，或运行 `1.0.入口.vb` 里的 `HelloVBA` 一章一章走

> 想让 `1.3` 里的 `ToDouble(...)` 能编译通过，必须先把 `1.6.工具函数.vb` 也粘进同一个工程。
> 漏粘会报「Sub 或 Function 未定义」——这不是 bug，见下文第 5 节。

---

## 2. 目录结构

```
VBA/
├── 1、类型/                 第 1 章：变量与数据类型
│   ├── 0.目录.vb            本章导读，只有注释，无代码
│   ├── 1.0.入口.vb          HelloVBA，总入口
│   ├── 1.1.基本数据类型.vb   9 种基础类型的范围与占用
│   ├── 1.2.简写形式.vb      类型声明字符、字面量后缀、$ 版函数
│   ├── 1.3.Variant与常量.vb  Variant / VarType / Const / 类型转换
│   ├── 1.4.数组.vb          固定、动态、二维数组与 Array/Split/Join/Filter
│   ├── 1.5.作用域.vb        Dim / Static / Private 三者的生命周期
│   ├── 1.6.工具函数.vb       ToDouble、ToFixed 两个可复用函数
│   └── 1.7.浮点与数值比较.vb 容差比较、Round 银行家舍入、精度常识
├── 2、流程控制/             第 2 章：判断与循环
│   ├── 0.目录.vb
│   ├── 2.0.入口.vb          HelloControl，总入口
│   ├── 2.1.条件判断.vb      If / IIf / Select Case、And / Or 不短路
│   ├── 2.2.For循环.vb       For...Next、Step、嵌套、For Each、Exit For
│   ├── 2.3.条件循环.vb      Do While / Do Until、前测与后测
│   └── 2.4.跳转与退出.vb    Exit Sub / GoTo 标签 / 卫语句
├── 3、过程与函数/           第 3 章：Sub / Function / 错误处理
│   ├── 0.目录.vb
│   ├── 3.0.入口.vb          HelloProcedures，总入口
│   ├── 3.1.Sub与Function.vb 返回值写法、调用写法、Public / Private
│   ├── 3.2.参数传递.vb      ByRef（默认）/ ByVal、Optional、ParamArray
│   ├── 3.3.错误处理.vb      On Error GoTo、Err 对象、Err.Raise
│   └── 3.4.递归与实用模式.vb 递归、Static、返回数组、命名冲突
├── 4、对象/                 第 4 章：Excel 对象模型
│   ├── 0.目录.vb
│   ├── 4.0.入口.vb          HelloObjects，总入口
│   ├── 4.1.对象基础.vb      Set / Nothing / Is / With / 默认属性
│   ├── 4.2.Sheet操作.vb     引用、遍历、增删、复制移动、隐藏、保护
│   ├── 4.3.单元格与区域.vb  Range / Cells / Offset / Resize / End / UsedRange
│   └── 4.4.表格ListObject.vb 超级表、结构化引用、增删行、筛选、排序、汇总
├── 5、实战基础/             第 5 章：文本 / 字典 / COM（量产必备）
│   ├── 0.目录.vb            本章导读 + 与 JS 的整体对照表
│   ├── 5.0.入口.vb          HelloPractical，总入口
│   ├── 5.1.字符串处理.vb    Trim / InStr / InStrRev / StrComp / Mid / Split /
│   │                        Join / Filter / Replace / Like / StrConv / 正则
│   ├── 5.2.文本文件与CSV.vb Open / Line Input / Print # / Dir / UTF-8 读写
│   ├── 5.3.Dictionary字典.vb CreateObject、Exists、映射与按 bin 统计
│   ├── 5.4.COM与CreateObject.vb ProgID、Excel.Application、释放防泄漏
│   └── 5.5.随机与模拟数据.vb Rnd / Randomize、造模拟数据、固定种子复现
├── CONTRIBUTING.md          Git 提交规范
└── README.md                本文件
```

章节顺序按依赖走：**类型 → 流程控制 → 过程与函数 → 对象 → 实战基础**。对象放倒数第二，是因为操作 `Range` 前得先会用循环遍历、用函数封装判定；第 5 章放最后，是因为它把前四章的能力组合成"读文件 → 切字符串 → 查表统计 → 出结果"的完整链路。

约定：目录名带序号，文件名 = 章节号 + 主题，`.vb` 后缀只是为了在编辑器里有高亮，实际是纯文本。

---

## 3. 章节索引

### 第 1 章 类型（已完成）

| 文件 | 主要知识点 | 关键结论 |
|---|---|---|
| `0.目录` | 章节导读、7 条「先记住这几条」 | 全局速查表，忘了就回来翻 |
| `1.0.入口` | `MsgBox`、单引号注释、`下划线` 续行 | 断点式入口，各节默认注释掉，逐节放开 |
| `1.1.基本数据类型` | `Byte` `Integer` `Long` `Single` `Double` `Currency` `String` `Boolean` `Date` `Object` | 循环计数器用 `Long`；金额用 `Currency`；`True = -1` |
| `1.2.简写形式` | `% & ^ ! # @ $` 声明字符与字面量后缀、`:` 一行多语句、`Trim$` 类函数、`DefType` | `30000 * 2` 会溢出，写 `30000& * 2`；文本处理优先用 `$` 版 |
| `1.3.Variant与常量` | 默认 `Variant`、`VarType` 返回值、`Const`、`CInt/CLng/CStr/CBool`、`IsNumeric` | 字符转浮点别用 `CDbl`，用 `Val` 或先校验再转 |
| `1.4.数组` | 一维/动态/二维、`ReDim Preserve`、`UBound/LBound`、`Array/Split/Join/Filter/Erase` | `Dim a(3)` 是 4 个元素；VBA 二维数组是矩形，不是数组套数组 |
| `1.5.作用域` | 过程级 `Dim`、静态 `Static`、模块级 `Private` / `Public` | 连按 F5 观察 `temp` 恒为 1、`keep` 一直累加 |
| `1.6.工具函数` | 按 `VarType` 分支的安全转换、`Err.Raise` 自定义错误、`Optional` 默认参数 | 脏数据抛错，绝不静默返回 0 |
| `1.7.浮点与数值比较` | `0.1+0.2<>0.3`、容差比较、`Round`/`CInt` 银行家舍入、`Int` vs `Fix`、累加误差、`Currency` | 判限用 `Abs(实测-规格) <= 容差`；`Round(2.5)=2`，要四舍五入用 `WorksheetFunction.Round` |

### 第 2 章 流程控制（已完成）

| 文件 | 主要知识点 | 关键结论 |
|---|---|---|
| `2.0.入口` | `HelloControl` 总入口 | 与第 1 章同样式，逐节放开 |
| `2.1.条件判断` | 块状/单行 `If`、`ElseIf`、`IIf`、`Select Case`（多值、`To` 区间、`Is`） | `And`/`Or` **不短路**；`IIf` 两个分支都会被求值，有副作用就别用 |
| `2.2.For循环` | `For...Next`、`Step`（含负数）、嵌套、`For Each`、`Exit For` | `Next` 之后计数器 = 终值 + 步长；`For Each` 变量必须是 `Variant` |
| `2.3.条件循环` | `Do While`/`Do Until`、前测与后测、`Exit Do`、`While...Wend` | 前测可能 0 次、后测至少 1 次；外部条件循环必须带次数上限 |
| `2.4.跳转与退出` | `Exit Sub`、`GoTo` 标签统一收尾、卫语句 | `GoTo` 只用于跳到末尾善后，日常分支用 `If`/`Select Case` |

### 第 3 章 过程与函数（已完成）

| 文件 | 主要知识点 | 关键结论 |
|---|---|---|
| `3.0.入口` | `HelloProcedures` 总入口 | — |
| `3.1.Sub与Function` | `Sub` 与 `Function` 的区别、返回值写法、`Call` 与括号规则、`Public`/`Private`、`Exit Function` | 返回值靠"给函数名赋值"；忘写就静默返回类型默认值 |
| `3.2.参数传递` | `ByRef`（默认）/ `ByVal`、`Optional` 默认值、`ParamArray` | 参数默认 **ByRef**，只读入参一律显式 `ByVal` |
| `3.3.错误处理` | `On Error GoTo`、`Err` 对象、`Resume`/`Resume Next`、`On Error GoTo 0`、`Err.Raise` | `Resume Next` 用完必须关掉，否则后面全在静默失败 |
| `3.4.递归与实用模式` | 递归（阶乘）、`Static` 记忆、返回数组、二义性名称 | `Long` 只装到 `12!`；返回数组必须声明成 `Variant` |

### 第 4 章 对象（已完成）

| 文件 | 主要知识点 | 关键结论 |
|---|---|---|
| `4.0.入口` | `HelloObjects` 总入口 | 会改数据的语句默认注释掉，避免按 F5 改坏工作簿 |
| `4.1.对象基础` | `Set`、`Nothing`、`Is`、`With`、默认属性、链式引用的取舍 | 对象赋值必须 `Set`；判断空用 `obj Is Nothing` |
| `4.2.Sheet操作` | 引用、遍历、增删、复制移动、隐藏、保护 | 遍历用 `For Each ws In ThisWorkbook.Worksheets` |
| `4.3.单元格与区域` | `Range`/`Cells`/`Offset`/`Resize`/`End`/`UsedRange`、`Value` 与 `Value2`、批量读写 | `Cells(行, 列)` 与 `Range("C2")` 的行列顺序相反；批量读写走数组 |
| `4.4.表格ListObject` | 超级表、结构化引用、增删行、筛选、排序、汇总 | `Find`/`AutoFilter` 会留下查找状态，记得复位 |

### 第 5 章 实战基础（已完成）

| 文件 | 主要知识点 | 关键结论 |
|---|---|---|
| `5.0.入口` | `HelloPractical` 总入口 | 只读写临时目录，不碰业务数据 |
| `5.1.字符串处理` | `Trim`/`LTrim`/`RTrim`、`InStr`/`InStrRev`/`StrComp`、`Left`/`Mid`/`Right`/`Mid` 语句、`Split`（`limit`、多分隔符）/`Join`/`Filter`、`Replace`、`UCase`/`LCase`/`StrConv`、`Asc`/`Chr`、`Like`、`CStr`/`Str`/`Val`/`Format`、`VBScript.RegExp`、拼接性能 | `Trim` 只去普通空格，全角空格/Tab 要自己换；`InStr` 找不到返回 0、`Split` 是 0 基；`InStrRev` 参数顺序与 `InStr` 相反；`Left$` 越界返回空串；循环拼接改用数组 + `Join` |
| `5.2.文本文件与CSV` | `Open`/`Line Input`/`Print #`/`FreeFile`/`EOF`、`Dir` 遍历、带引号 CSV 解析、`ADODB.Stream` 读写 UTF-8 | 原生 `Open` 是 ANSI，中文 UTF-8 必须走 `ADODB.Stream`；`Print #` 的逗号是制表位 |
| `5.3.Dictionary字典` | `CreateObject`、`Exists`、`Keys`/`Items`、`CompareMode`、映射与计数 | 取值前必须 `Exists`，直接 `d(新键)` 会新建空条目 |
| `5.4.COM与CreateObject` | 后期绑定、常用 ProgID、`Excel.Application`、释放与防泄漏 | 进程外对象要 `Quit` + `Set Nothing`，顺序反了留僵尸进程 |
| `5.5.随机与模拟数据` | `Rnd`/`Randomize`、整数区间公式、标称±容差造数据、固定种子 | 不播种则每次同一条序列；`Rnd(0)` 不前进；正式量产结果不能用随机 |

---

## 4. 核心知识点速查

**类型与内存**

| 类型 | 字节 | 范围 / 说明 |
|---|---|---|
| `Byte` | 1 | 0 ~ 255 |
| `Integer` | 2 | -32,768 ~ 32,767（**上限只有 32767**） |
| `Long` | 4 | ±21 亿，计数器首选 |
| `Single` | 4 | 约 7 位有效数字 |
| `Double` | 8 | 约 15 位有效数字，工程计算首选 |
| `Currency` | 8 | 定点 4 位小数，无浮点误差，金额专用 |
| `Boolean` | 2 | `True = -1`，`False = 0` |
| `Date` | 8 | 100/1/1 ~ 9999/12/31 |
| `String` | 变长 | 变长字符串 |
| `Variant` | 16+ | 不写 `As` 时的默认类型，最灵活也最慢 |

**常用转换函数**

- 字符 → 整数：`CInt` / `CLng`（会取整，`.5` 走银行家舍入：`CInt(22.5) = 22`）
- 字符 → 浮点：`Val`（固定按英文 `.` 解析，不看区域设置）或 `CDec`
- 取整：`Fix`（向 0 取整）与 `Int`（向下取整），负数时结果不同
- 判断：`IsNumeric` / `IsEmpty` / `IsArray` / `IsNull` / `VarType`

**数组三件套**

```vb
Dim a(1 To 3) As Integer      ' 固定：1~3，共 3 个
Dim b() As String            ' 动态：先不写长度
ReDim Preserve b(1 To 3)     ' 保留已有数据的扩容
```

**流程控制与过程**

```vb
If x > hi Then Exit Sub                    ' 单行 If；卫语句，提前走人
Select Case n
    Case 1 To 5, 8                         ' 区间 + 枚举，逗号分隔
    Case Is > 10                           ' Is 后只能跟比较运算符
    Case Else
End Select

For i = 1 To 10 Step 2                     ' Next 后 i = 12
Do While cond : Loop                       ' 前测，可能 0 次
Do : Loop While cond                       ' 后测，至少 1 次

Sub Foo(ByVal v As Double)                 ' 只读入参写 ByVal（默认是 ByRef）
Function Bar(v As Double) As Double
    Bar = v * 2                            ' 返回值 = 给函数名赋值
End Function
```

---

## 5. 本机实测踩坑清单

这些是实际敲代码时炸出来的坑，不是抄文档抄来的，遇到同类问题优先回这里查。

| # | 现象 | 原因 / 解法 |
|---|---|---|
| 1 | `CDbl("22.5")` 报「运行时错误 6：溢出」，连手写常量都炸 | 系统区域设置里「小数点符号」不是英文句点。属环境问题，不是代码问题。字符串转浮点统一用 `Val`，或走 `1.6` 的 `ToDouble` |
| 2 | `30000 * 2` 报「溢出」 | 无后缀整数字面量默认按 `Integer` 运算，结果超 32767。写成 `30000& * 2` 强制 `Long` |
| 3 | `Dim a(3)` 数出 4 个元素，逻辑越界 | 省略下界时默认从 0 开始，`0~3` 共 4 个。要精确控制就写 `1 To 3` |
| 4 | `Dim a, b As Integer` 只有 `b` 是整数 | `As` 只作用于紧邻的那个变量，`a` 仍是 `Variant`。要都声明就逐个写 `As` |
| 5 | `arr2 = arr1` 编译报错 | VBA 不支持数组整体赋值，必须逐元素拷；或两边都用 `Variant` 装 |
| 6 | `Erase` 之后再 `UBound` 报「下标越界」 | `Erase` 把动态数组打回「未分配」状态，必须重新 `ReDim` |
| 7 | 模块名带点，粘贴后无法命名 | VBA 模块名不允许 `.`，文件名 `1.4.数组.vb` 对应模块名写 `1_4_数组` |
| 8 | `LongLong` 报「用户定义类型未定义」 | `LongLong` / `^` 声明字符仅 64 位 Office 支持 |
| 9 | `IsNumeric(日期)` 返回 `True` | VBA 把日期当数字，做数值校验时要显式排除 `vbDate`（见 `ToDouble`） |
| 10 | `If Not obj Is Nothing And obj.Row > 0` 照样报错 | `And` / `Or` **不短路**，两边都会算。判空与取值必须拆成两层 `If` |
| 11 | 调用函数后发现外面变量的值被改了 | 参数默认是 `ByRef`，函数里的修改会写回调用方。只读入参一律加 `ByVal` |
| 12 | `On Error Resume Next` 之后的代码全在静默失败 | 忘了 `On Error GoTo 0` 关闭。用完立刻关，且每轮循环先 `Err.Clear` |
| 13 | `For Each` 遍历数组报「需要对象」 | 循环变量必须声明成 `Variant`（或对象类型），不能写 `String` / `Long` |
| 14 | 函数明明该返回 `True` 却总是 `False` | 忘了写 `函数名 = ...`。VBA 不报错，直接返回类型默认值（`False` / `0` / `""`） |

**关于「VBA 没有 import」**——这是从 JS 转过来最容易困惑的点，值得单拎出来：

- 一个 `.xlsm` = 一个 `VBAProject`，工程内所有模块（标准模块、类模块、窗体、`ThisWorkbook`、`Sheet1`）共享同一个全局命名空间
- `Sub` / `Function` 不写修饰符默认就是 `Public`，编译期一次性解析全部名字，直接按名字调用即可
- 报「Sub 或 Function 未定义」= 对应模块没粘进工程，这才是 JS 里「忘记 import」的等价场景
- 想限制成只给本模块用，加 `Private`；想整个模块对外不可见，模块首行加 `Option Private Module`
- 代价是**重名即冲突**（报「二义性名称」），过程名别用 `Range`、`Date` 这类内置名
- 真正需要「引用」的只有跨工程：`工具 → 引用` 勾选对象库，或用 `Application.Run "别的文件.xlsm!宏名"`

---

## 6. 编码与命名约定

- **文件编码统一 UTF-8**。注释里有中文，某台机器 VBA 编辑器显示乱码时先查编码，别把乱码提交上去
- **一个 `.vb` 文件 = 一个标准模块**，粘进 VBA 后模块名改写为下划线版本
- 所有模块首行加 `Option Explicit`，强制声明变量；新增模块务必别漏
- 变量名用驼峰（`grid`、`pins`），常量全大写（`MAX_RETRY`、`VOLT_TOL`）
- 注释用单引号 `'`，章节标题用 `' ====` 分隔线，保持现有文件风格

---

## 7. 提交规范

完整规则见 [CONTRIBUTING.md](./CONTRIBUTING.md)，核心是 Conventional Commits：

```
feat(类型): 新增 1.4 数组小节，含 ReDim Preserve 示例
fix(工具函数): ToDouble 参数改为 Variant，避免区域设置把 3.14 转成 314
docs(类型): 补充 VBA 无需 import 的说明
```

要点：

- `scope` 写章节号或模块名（`类型`、`1.4`、`工具函数`）
- `subject` 用中文、动词开头、不加句号、50 字内
- 一次提交只做一件事；`body` 重点写**为什么这么改**和踩到的坑
- Excel 工作簿不入库，`.gitignore` 应包含 `~$*.xls*`、`*.xlsm`、`*.xlsb`
- 主干保持可运行，日常学习走 `learn/<章节>` 分支

---

## 8. 后续规划

1. 第 6 章「完整流程」：CSV 原始数据 → 判限（含容差）→ 良率统计 → 生成报告
2. 第 3 章待补：类模块（`Class Module`）、`Property Let/Get/Set`、自定义对象
3. 第 4 章待补：`Workbook` 打开/另存/关闭、事件（`Worksheet_Change`、`Workbook_Open`）
4. 抽一个公共模块 `Util.bas`，把 `ToDouble`、`ToFixed`、`LastRowOf`、`Judge`、`SplitCsvLine`、`ReadUtf8`，以及 5.1 末尾的 `StartsWith`/`EndsWith`/`PadLeft`/`PadRight`/`TrimChars`/`SplitAny` 收进去，避免各章重复定义触发「二义性名称」
5. 第 5 章待补：`FSO` 递归遍历目录、`ADODB` 连数据库（正则已并入 5.1 的 `DemoRegex`）
6. 全章示例统一往 ATE 场景靠：引脚电压、上下限判定、良率统计
