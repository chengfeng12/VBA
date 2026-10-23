Option Explicit

' ============================================================
' 2.2 For 循环：循环次数已知时用它
'     For 计数变量 = 初值 To 终值 [Step 步长]
'         ...
'     Next [计数变量]
'     步长省略时默认 1；可以是负数（倒序）
' ============================================================
Sub DemoFor()
    Dim i As Long, sum As Long

    ' ---- 最基本的形式：1 到 3，循环 3 次 ----
    For i = 1 To 3
        sum = sum + i
    Next
    MsgBox "1+2+3 = " & sum & "，循环结束后 i = " & i
    ' ★ Next 之后 i = 终值 + 步长（这里是 4，不是 3）。
    '   靠这个特性做"多走一格"的计算时，一定在注释里写清楚

    ' ---- Step 步长 ----
    Dim s As String
    For i = 1 To 10 Step 2
        s = s & i & " "
    Next
    MsgBox "步长 2 → " & s

    s = ""
    For i = 3 To 1 Step -1
        s = s & i & " "
    Next
    MsgBox "倒序 → " & s
    ' ★ Step 0 会死循环（计数器永远停在初值），别写

    ' ---- 嵌套循环：外层行、内层列，遍历二维数组（1.4 的 grid）----
    Dim grid(1 To 2, 1 To 3) As Double
    Dim r As Long, c As Long
    grid(1, 1) = 3.3:  grid(1, 2) = 3.31: grid(1, 3) = 3.29
    grid(2, 1) = 0#:   grid(2, 2) = 0.01: grid(2, 3) = 0#

    For r = 1 To 2
        For c = 1 To 3
            grid(r, c) = grid(r, c) * 1000        ' 单位 V → mV
        Next c                                     ' Next 后面写变量名，嵌套时不会错配
    Next r

    MsgBox "grid(1,2) = " & grid(1, 2) & " mV"

    ' ---- Exit For：提前跳出，只跳出"当前这一层" ----
    Dim found As Long
    For i = 1 To 100
        If i * i > 500 Then
            found = i
            Exit For
        End If
    Next
    MsgBox "第一个平方大于 500 的数是 " & found
    ' ★ 嵌套里 Exit For 只退出内层，外层会继续；要一起退就加标志位或抽成函数（3.1）

    ' ---- For Each：不关心下标时更省事 ----
    Dim codes As Variant, p As Variant, txt As String
    codes = Array("VCC", "GND", "CLK")
    For Each p In codes
        txt = txt & p & " "
    Next
    MsgBox "For Each → " & txt
    ' ★ 坑：For Each 的循环变量必须是 Variant（或对象类型），
    '   写成 Dim p As String 会报"需要对象"，因为 VBA 不知道数组元素的真实类型
    ' ★ For Each 只适合"只读遍历"：要改元素值还得用下标 For
    ' ★ 集合（Worksheets / ListObjects）也能 For Each，变量写成对应对象类型，见第 4 章
End Sub
