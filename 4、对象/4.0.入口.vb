Option Explicit

' ============================================================
' 4.0 入口：本章总入口
'     下面几行调用默认注释掉了，去掉注释即可依次运行
'     也可以把光标放进任意一个 Sub 里按 F5 单独运行
'     ★ 本章多数语句会真实修改工作簿，"会改数据"的代码都写在注释里
' ============================================================
Sub HelloObjects()
    MsgBox "第 4 章：对象 —— 表格 / 单元格 / Sheet"

    ' 本章各小节（每个 Sub 在独立模块里）
    ' DemoObjectBasics     ' 4.1 对象基础：Set / Nothing / With
    ' DemoSheets           ' 4.2 Sheet 操作
    ' DemoRange            ' 4.3 单元格与区域
    ' DemoTable            ' 4.4 表格（ListObject）
End Sub
