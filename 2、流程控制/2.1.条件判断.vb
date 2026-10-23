Option Explicit

' ============================================================
' 2.1 条件判断：If / ElseIf / Else、IIf、Select Case
'     VBA 的 If 有两种写法：
'       块状（多行，必须 End If）
'       单行（一行写完，不能写 End If，也不能带 ElseIf）
' ============================================================
Sub DemoIf()
    Dim v As Double, lo As Double, hi As Double
    v = 3.3
    lo = 3#
    hi = 3.6

    ' ---- 块状 If：分支里有多个语句时必须用这种 ----
    If v < lo Then
        MsgBox "偏低：" & v & " < " & lo
        MsgBox "判定：Fail"
    ElseIf v > hi Then
        MsgBox "偏高：" & v & " > " & hi
        MsgBox "判定：Fail"
    Else
        MsgBox "合格：" & lo & " ≤ " & v & " ≤ " & hi
    End If
    ' 条件从窄到宽 / 从特殊到一般，命中一个分支就走，后面的不再判断

    ' ---- 单行 If：分支只有一句时的简写，一行内写完 ----
    If v > hi Then MsgBox "超上限"
    If v > hi Then MsgBox "超上限" Else MsgBox "未超上限"
    ' ★ 单行写法不能有 End If，也不能写 ElseIf；
    '   分支一多就退回块状写法，别硬挤成一行

    ' ---- 逻辑运算符：And / Or / Not / Xor ----
    If v >= lo And v <= hi Then MsgBox "区间内"
    If v < lo Or v > hi Then MsgBox "越界"
    If Not (v > hi) Then MsgBox "没有超上限"

    ' ★ 最大的坑：VBA 的 And / Or 没有短路求值，两边一定都会算完
    '   JS 里 a && b 在 a 为假时不看 b；VBA 一定把 b 也算一遍，
    '   所以不能靠"前面挡住后面"：先判空再取值，必须拆成两层 If，不能靠 And 串起来
    '   （第 4 章判单元格/对象是否存在时尤其明显）

    ' ---- 三元运算：VBA 没有 ?: ，用 IIf 函数 ----
    Dim result As String
    result = IIf(v >= lo And v <= hi, "Pass", "Fail")
    MsgBox "IIf → " & result

    ' ★ 坑：IIf 是"函数"，不是"语句"——两个分支（包括真假两侧的表达式）
    '   都会先算出来，再挑一个返回。所以 IIf(x <> 0, 1 / x, 0) 在 x = 0 时
    '   照样会除零报错。只要有副作用或可能出错，就老老实实写 If

    ' ---- Select Case：多分支比一长串 ElseIf 清楚 ----
    Dim code As String
    code = " pass "
    Select Case UCase$(Trim$(code))            ' Select 后面只能跟一个表达式
        Case "PASS", "OK"                      ' 一个 Case 可以跟多个值，逗号分隔
            MsgBox "判定为合格"
        Case "FAIL", "NG"
            MsgBox "判定为不合格"
        Case Else                              ' 兜底分支，放最后
            MsgBox "未知结果：[" & code & "]"
    End Select

    ' ---- Case 的三种写法 ----
    Dim n As Long
    n = 7
    Select Case n
        Case 1 To 5                            ' 闭区间
            MsgBox "1~5"
        Case 6, 8                              ' 枚举
            MsgBox "6 或 8"
        Case Is > 10                           ' Is 后面只能跟 > >= < <= = <>
            MsgBox "大于 10"
        Case Else
            MsgBox "其他：" & n
    End Select
    ' ★ 从上往下匹配，命中第一个就走 → 范围有重叠时，窄的（特殊的）写前面
    ' ★ 要同时判断两个变量（如 同时看电压和电流），Select Case 做不到，用 If
End Sub
