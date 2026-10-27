Option Explicit

' ============================================================
' 第 1 章  类型（变量与数据类型）
' ------------------------------------------------------------
' 用法：
'   每个 .vb 文件 = 一个独立的 VBA 标准模块
'   Excel → Alt+F11 → 插入 → 模块 → 把整个文件内容粘进去
'   模块名建议和文件名对应（VBA 模块名不能带点，可用 1_1_基本数据类型）
' ------------------------------------------------------------
' 目录
'   1.0.入口.vb            HelloVBA —— 本章总入口，一次性跑完全部演示
'   1.1.基本数据类型.vb    Byte / Integer / Long / Single / Double / Currency
'                          / String / Boolean / Date / Object 的范围与占用
'   1.2.简写形式.vb        类型声明字符 % & ^ ! # @ $、字面量后缀、DefInt、$ 版函数
'   1.3.Variant与常量.vb   Variant、VarType、Const、类型转换（含 CDbl 避坑）
'   1.4.数组.vb            一维 / 动态 / 二维数组、ReDim / ReDim Preserve
'                          （扩容、缩容、只能改最后一维、O(n²) 扩容）、
'                          LBound / UBound（元素个数、维号）
'                          Array、Split、Join、Filter、Erase
'                          常用操作：push/pop、查找、切片、合并、排序、
'                          条件过滤、一维↔二维 Transpose
'   1.5.作用域.vb          Dim / Static / Private 三种作用域的区别
'   1.6.工具函数.vb        ToDouble —— 字符串安全转 Double
'   1.7.浮点与数值比较.vb  0.1+0.2<>0.3、容差比较、Round 银行家舍入、
'                          Int 与 Fix、精度常识
' ------------------------------------------------------------
' 先记住这几条
'   1. Dim 不写 As 就是 Variant：最灵活也最慢，正式代码尽量写明类型
'   2. Dim i%, l&, s!, d#, c@, str$ 是类型声明字符的简写
'   3. Dim a(3) 是 0~3 共 4 个元素，不是 3 个（ReDim a(3) 同理）
'   4. Dim a, b As Integer 只有 b 是 Integer，a 是 Variant
'   5. 本机 CDbl("22.5") 会报「运行时错误 6：溢出」，字符串转浮点统一用 Val
'   6. Integer 上限只有 32767，循环计数器一律用 Long
'   7. VBA 没有 import：一个 .xlsm = 一个工程，工程内所有模块共享同一个
'      命名空间，Sub/Function 默认就是 Public，直接按名字互相调用；
'      只有"模块没粘进工程"才会在编译期报「Sub 或 Function 未定义」。
'      代价是重名即冲突（报"二义性名称"），所以过程名别用 Range/Date 这类内置名
'   8. 两个 Double 永远别用 = 比：0.1+0.2 <> 0.3，判限要用 Abs(a-b) <= 容差
'   9. Round / CInt 是银行家舍入（Round(2.5)=2），要四舍五入用
'      Application.WorksheetFunction.Round
' ============================================================

' 本模块只放说明，没有可执行的 Sub
