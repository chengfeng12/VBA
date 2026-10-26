Option Explicit

' ============================================================
' 第 4 章  对象（Excel 表格 / 单元格 / Sheet 操作）
' ------------------------------------------------------------
' 用法：每个 .vb 文件 = 一个独立的 VBA 标准模块，粘进 Excel 工程即可
'       本章直接操作真实工作簿，所以"会改数据 / 改结构"的语句一律
'       写在注释里，需要时自己取消注释（避免一按 F5 就把表改坏）
' ------------------------------------------------------------
' 目录
'   4.0.入口.vb              HelloObjects —— 本章总入口
'   4.1.对象基础.vb          Set / Nothing / Is / With / 默认属性、
'                            TypeOf、Dim As New 陷阱、Collection / Dictionary
'   4.2.Sheet操作.vb         引用、遍历、增删、复制移动、隐藏、保护、
'                            CodeName、提速三件套
'   4.3.单元格与区域.vb      Range / Cells / Offset / Resize / End / UsedRange
'                            Value 与 Value2、批量读写、Find、合并单元格
'   4.4.表格ListObject.vb    超级表、结构化引用、增删行、筛选、排序、汇总、
'                            空表 DataBodyRange 陷阱
' ------------------------------------------------------------
' 对象模型（从上往下是父子关系）
'
'   Application                    整个 Excel 程序
'    ├─ ScreenUpdating / Calculation 全局开关（提速用）
'    ├─ Workbooks                   所有打开的工作簿
'    │   └─ Workbook                一个 .xlsm 文件
'    │       ├─ Worksheets          普通工作表集合
'    │       ├─ Sheets              全部表（含图表工作表、宏表）
'    │       └─ Names / Styles ...
'    └─ ...
'
'   Worksheet                      一张工作表
'    ├─ Range / Cells               单元格、区域（最常用）
'    ├─ ListObjects                 超级表（结构化引用）
'    ├─ UsedRange                   用过的范围
'    └─ Shapes / ChartObjects ...
'
'   三个"当前"，务必分清：
'     ThisWorkbook    代码所在的工作簿  ← 最稳，优先用这个
'     ActiveWorkbook  当前活动的工作簿  ← 用户一切换就变，别依赖
'     ActiveSheet     当前活动的表      ← 同上
' ------------------------------------------------------------
' 必记的几条
'   1. 对象变量赋值必须用 Set；对象之间赋值是"引用"，不是拷贝
'   2. 用完 Set obj = Nothing 释放引用；判断空用 obj Is Nothing（不是 =）
'   3. Cells(行, 列) 的参数顺序是 (row, column)，而 Range("A1") 是"先列后行"
'   4. 批量读写一定走数组：arr = Range("A1:C100").Value，比逐格快几十倍
'   5. 提速三件套：ScreenUpdating / Calculation / EnableEvents，用完必须恢复
'   6. Find / AutoFilter 之类会改"查找状态"，记得复位，否则影响用户后续操作
'   7. 超级表只有表头时 DataBodyRange 是 Nothing，读整表前必须先判
'   8. Find 默认从区域首格的"下一个"开始找，要扫全区域得传 After:=最后一格
'   9. Dim x As New 会让 x Is Nothing 永远为 False，要延迟建就写 Set x = New
' ============================================================

' 本模块只放说明，没有可执行的 Sub
