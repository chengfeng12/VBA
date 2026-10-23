Option Explicit

' ============================================================
' 1.5 作用域：变量在哪一层声明，决定谁能用它
'     Dim     —— 过程级，过程结束就销毁
'     Static  —— 过程级，但值在多次调用之间保留
'     Private —— 模块级，本模块所有过程共享（Public 则整个工程可见）
' ============================================================
Private gCounter As Long             ' 模块级：本模块所有 Sub 都能用

Sub DemoScope()
    Dim temp As Long                 ' 过程级：只在 DemoScope 里有效
    Static keep As Long              ' 过程级，但调用结束后值会保留

    temp = temp + 1                  ' 每次调用都从 0 开始，永远是 1
    keep = keep + 1                  ' 每次调用累加，用来做"调用次数"计数
    gCounter = gCounter + 1          ' 模块级，跨 Sub 共享

    MsgBox "temp = " & temp & vbCrLf & _
           "keep = " & keep & vbCrLf & _
           "gCounter = " & gCounter & vbCrLf & vbCrLf & _
           "对着这个 Sub 连按 F5，temp 永远是 1，keep 和 gCounter 每次加 1"
End Sub
