Option Explicit

' ============================================================
' 3.2 参数传递：ByRef（默认）与 ByVal
'     ★ VBA 的参数默认是 ByRef（传地址）：函数里改了，外面的变量跟着变
'       这和 JS / Python 的直觉不一样，是从别的语言转过来最容易踩的坑
'     ★ 规则：只读取的入参一律写 ByVal，需要往回带结果才用 ByRef
' ============================================================
Sub DemoPassing()
    Dim v As Double
    v = 3.3

    ' ---- 默认 ByRef：函数内部改了，外面的 v 跟着变 ----
    ClampRef v, 3#, 3.6
    MsgBox "ClampRef 之后 v = " & v            ' 3.3 → 3.3（本来就合格）

    v = 3.9
    ClampRef v, 3#, 3.6
    MsgBox "ClampRef 之后 v = " & v            ' 3.9 被改回 3.6

    v = 3.9
    ClampVal v, 3#, 3.6
    MsgBox "ClampVal 之后 v = " & v            ' 还是 3.9，ByVal 保护了入参

    ' ---- Optional：可选参数，必须放在参数表最后 ----
    MsgBox "默认 4 位小数：" & ToFixedDemo(3.14159265)
    MsgBox "指定 2 位小数：" & ToFixedDemo(3.14159265, 2)
    ' 默认值只能是常量（3.4 里用 Variant + IsMissing 可以做"判断有没有传"）

    ' ---- ParamArray：参数个数不定 ----
    MsgBox "求和 1+2+3+4+5 = " & SumAll(1, 2, 3, 4, 5)
    MsgBox "平均三个采样点 = " & AvgAll(3.3, 3.31, 3.29)
    ' ★ ParamArray 只能放在最后一个参数，且只能声明成 Variant 数组；
    '   传进来后在函数里就是一个下标从 0 开始的数组

    ' ---- 对象参数：传的永远是引用 ----
    ' 第 4 章：Sub Foo(ws As Worksheet) 里改的单元格，就是调用方那张表，
    ' ByVal 对对象也拦不住"改对象内部状态"，只能拦住"换指向"
End Sub


' 默认 ByRef：把入参就地改成合法范围内的值，调用方能看到修改
Sub ClampRef(v As Double, ByVal lo As Double, ByVal hi As Double)
    If v < lo Then v = lo
    If v > hi Then v = hi
End Sub


' 显式 ByVal：只在本过程里改，外部变量不受影响
Sub ClampVal(ByVal v As Double, ByVal lo As Double, ByVal hi As Double)
    If v < lo Then v = lo
    If v > hi Then v = hi
    MsgBox "ByVal 内部已改成 " & v & "，但外面拿不到这个结果"
End Sub


' Optional + 默认值
Function ToFixedDemo(ByVal d As Double, Optional ByVal digits As Long = 4) As String
    ToFixedDemo = Format$(d, "0." & String$(digits, "0"))
    ' 注意：Format 按系统区域设置输出，小数点不是英文句点的机器上要改用 Replace
End Function


' ParamArray：不限个数的数值求和
Function SumAll(ParamArray nums() As Variant) As Double
    Dim i As Long
    For i = LBound(nums) To UBound(nums)
        SumAll = SumAll + nums(i)
    Next
End Function


' 一个参数都没传时，UBound 是 -1，得先挡一下再除
Function AvgAll(ParamArray nums() As Variant) As Double
    Dim i As Long, cnt As Long
    If UBound(nums) < LBound(nums) Then Exit Function
    For i = LBound(nums) To UBound(nums)
        AvgAll = AvgAll + nums(i)
        cnt = cnt + 1
    Next
    AvgAll = AvgAll / cnt
End Function
