Option Explicit

' ============================================================
' 3.0 入口：本章总入口
'     下面几行调用默认注释掉了，去掉注释即可依次运行
'     也可以把光标放进任意一个 Sub 里按 F5 单独运行
' ============================================================
Sub HelloProcedures()
    MsgBox "第 3 章：过程与函数 —— Sub / Function"

    ' 本章各小节（每个 Sub 在独立模块里）
    ' DemoProcBasics    ' 3.1 Sub 与 Function
    ' DemoPassing       ' 3.2 参数传递 ByRef / ByVal / Optional
    ' DemoErrorHandling ' 3.3 错误处理
    ' DemoCatchCustomError ' 3.3 接住自定义错误（Err.Raise 的另一半）
    ' DemoRecursion     ' 3.4 递归与实用模式
End Sub
