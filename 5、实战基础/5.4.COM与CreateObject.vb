Option Explicit

' ============================================================
' 5.4 COM 与 CreateObject：机台脚本、跨应用自动化的地基
'     一句话：COM 就是"给系统一个名字，让它帮你启动或连上一个可编程对象"
'     VBT / VBScript / VBA 里都是这一套语法，学一次三处通用
' ------------------------------------------------------------
' 和 JS 对照：
'     require("fs")                      →  CreateObject("Scripting.FileSystemObject")
'     new Excel.Application()            →  CreateObject("Excel.Application")
'     连接已在运行的实例                    →  GetObject(, "Excel.Application")
'     手动 free / dispose                →  Set obj = Nothing
' ============================================================
Sub DemoComBasics()
    ' ---- 后期绑定：不勾引用，按 ProgID 字符串创建 ----
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    MsgBox "临时目录：" & fso.GetSpecialFolder(2).Path
    Set fso = Nothing

    ' 早期绑定写法（需勾引用，有属性提示，换机器可能得重新勾）：
    ' Dim fso2 As Scripting.FileSystemObject
    ' Set fso2 = New Scripting.FileSystemObject

    ' ---- 常用 ProgID ----
    ' Scripting.FileSystemObject   文件/目录操作（能递归、能判断存在，比 Dir 强）
    ' Scripting.Dictionary         字典（见 5.3）
    ' WScript.Shell                执行外部程序、读环境变量
    ' ADODB.Stream                 按指定编码读写文本（5.2 的 UTF-8 方案）
    ' Excel.Application            Excel 本体（见下一个过程）
    ' ★ ProgID 拼错不会编译报错，只在运行时炸（429：ActiveX 部件不能创建对象）

    ' ---- 释放 ----
    ' Set x = Nothing 是"松手"：引用计数归零，对象才被回收。
    ' 落在过程里的小对象不写也能被回收，
    ' 但"进程外"的对象（Excel.Application、Word.Application）不松手就会一直挂着。
End Sub


' ============================================================
' Excel.Application：重量级方案，量产上要克制
'   ★ 优先用 CSV + 文本流（5.2）：不弹窗、不吃内存、不受用户干扰；
'     只有"必须算公式 / 必须读 xlsm 结构 / 必须出图"时才动 COM
' ============================================================
Sub DemoExcelCom()
    Dim app As Object, wb As Object
    Set app = CreateObject("Excel.Application")     ' 新建一个隐藏的 Excel 进程
    app.Visible = False
    app.DisplayAlerts = False

    On Error GoTo Cleanup
    ' ★ 真实场景这么写（路径按需改）：
    ' Set wb = app.Workbooks.Open(ThisWorkbook.Path & "\data.xlsm")
    ' MsgBox "读到：" & wb.Worksheets(1).Range("A1").Value
    ' wb.Close SaveChanges:=False
    MsgBox "已启动一个隐藏的 Excel 进程，接着会立刻退出它"

Cleanup:
    On Error Resume Next
    app.DisplayAlerts = True
    app.Quit                                  ' ★ 先 Quit 再松引用，顺序反了会留僵尸进程
    Set wb = Nothing
    Set app = Nothing
    ' ★ 典型症状：任务管理器里一堆 EXCEL.EXE 越堆越多，
    '   基本都是"创建了没 Quit"或"Quit 之后还留着子对象引用"
End Sub


' ============================================================
' 几个实战里绕不开的坑
' ============================================================
Sub DemoComTraps()
    ' ---- 1) 429 无法创建对象 ----
    ' 常见原因：组件没注册（绿色版 Office）、32/64 位不匹配、
    '           权限不足、被杀软拦。先换一个 ProgID 试，能区分"环境问题"还是"这个对象的问题"

    ' ---- 2) GetObject：连上已经开着的实例 ----
    ' Set app = GetObject(, "Excel.Application")   ' 注意前面有个空参数
    ' ★ 没有已运行的实例时会报错 429；想"有就复用、没有就新建"，
    '   得用 On Error Resume Next 探一下再决定

    ' ---- 3) 跨进程调用很慢 ----
    ' 每个 .Range("A1").Value 都是一次进程间通信；
    ' 循环里逐格读一个表 = 几千次通信，改成数组一次搬（见 4.3）

    ' ---- 4) 服务/无人值守环境 ----
    ' 机台常以服务账号或计划任务跑脚本，此时没有交互桌面，
    ' Office COM 可能直接失败或卡住不返回 —— 这也是"量产优先文本流"的原因

    ' ---- 5) WScript.Shell 调外部程序 ----
    ' Dim sh As Object
    ' Set sh = CreateObject("WScript.Shell")
    ' sh.Run "cmd /c dir", 1, True      ' 第三个参数 True = 等它跑完
    ' Set sh = Nothing
End Sub
