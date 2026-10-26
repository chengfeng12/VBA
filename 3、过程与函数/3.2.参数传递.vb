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

    ' ---- ★ 加一对括号 = 强制按值传，修改静默丢失 ----
    v = 3.9
    ClampRef (v), 3#, 3.6                      ' 注意这个括号
    MsgBox "ClampRef (v) 之后 v = " & v        ' 还是 3.9！
    ' 原因：括号把 v 变成"表达式"，VBA 先求值、传的是一份临时副本，
    '       过程里改的是副本，调用方拿不到。
    ' JS 没有这种语法，从 JS 过来的人最容易在这里丢结果，而且不报错
    ' 同理：直接传字面量 ClampRef 3.9, 3#, 3.6 也合法，改的也是临时值

    ' ---- Optional：可选参数，必须放在参数表最后 ----
    MsgBox "默认 4 位小数：" & ToFixedDemo(3.14159265)
    MsgBox "指定 2 位小数：" & ToFixedDemo(3.14159265, 2)
    ' 默认值只能是常量，不能写 Date() 这类函数调用（和 JS 的 = 默认参数不同）

    ' ---- 判断"到底传没传"：只有 Variant 参数能用 IsMissing ----
    MsgBox "没传标签：" & Describe(3.3)
    MsgBox "传了标签：" & Describe(3.3, "VCC")
    ' Optional digits As Long = 4 这种参数拿不到"是否传了"，
    ' IsMissing 对它恒为 False，只能靠默认值兜底

    ' ---- 命名参数：按名字传，顺序可以打乱 ----
    MsgBox "命名参数：" & ToFixedDemo(digits:=2, d:=3.14159)
    ' JS 里靠对象解构 { digits: 2 } 达到同样效果，VBA 直接写"名字:="
    ' 参数多的函数建议两边都写名字，读代码时不用回翻参数表

    ' ---- ParamArray：参数个数不定 ----
    MsgBox "求和 1+2+3+4+5 = " & SumAll(1, 2, 3, 4, 5)
    MsgBox "平均三个采样点 = " & AvgAll(3.3, 3.31, 3.29)
    ' ★ ParamArray 只能放在最后一个参数，且只能声明成 Variant 数组；
    '   传进来后在函数里就是一个下标从 0 开始的数组

    ' ---- 对象参数：传的永远是引用 ----
    ' 第 4 章：Sub Foo(ws As Worksheet) 里改的单元格，就是调用方那张表，
    ' ByVal 对对象也拦不住"改对象内部状态"，只能拦住"换指向"

    ' ---- 数组参数：默认 ByRef 不复制，但调用时也不能加括号 ----
    '   Sub Fill(arr() As Double)      ' 数组形参写成 arr() As T
    '   Fill a                         ' 正确：只给数组名
    '   Fill (a)                       ' 错误：括号让数组变成表达式，报类型不匹配
    ' 几万个元素的数组按值传会明显变慢，大数组保持默认的 ByRef
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


' Optional + IsMissing：只有 Variant 形参才能区分"没传"和"传了空值"
' （JS 里可以写 function f(tag = undefined) 再判断，VBA 只能靠 IsMissing）
Function Describe(ByVal v As Double, Optional ByVal tag As Variant) As String
    If IsMissing(tag) Then
        Describe = v & " V（未命名）"
    Else
        Describe = tag & " = " & v & " V"
    End If
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
