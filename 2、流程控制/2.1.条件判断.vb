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

    ' ---- Not：唯一的一元运算符，但它其实是"按位取反"，不是纯逻辑非 ----
    Dim flag As Boolean
    flag = (v > hi)                                ' 3.3 > 3.6 → False
    If Not flag Then MsgBox "Not False = True"     ' 成立，走进来

    MsgBox "Not True  = " & (Not True)             ' False
    MsgBox "Not 0     = " & (Not 0)                ' -1   ← 不是 True
    MsgBox "Not -1    = " & (Not -1)               ' 0
    MsgBox "Not 5     = " & (Not 5)                ' -6   ← 按位取反，Not x = -x - 1
    ' ★ VBA 里 True 的数值就是 -1，根源正是 Not 0 = -1；
    '   JS 的 ! 只返回布尔值（!0 === true），VBA 的 Not 作用在整数上是逐位取反，
    '   所以 Not 只用来取反"布尔标志/比较结果"，别拿它当 ! 用在数字上

    ' ★ 优先级：比较运算符 > Not > And > Or，所以
    '   If Not v > hi Then          ' 等于 Not (v > hi)，比较先算
    '   而 JS 里 !v > hi 是 (!v) > hi —— 语义正好相反，跨语言时务必加括号

    ' ★ 对象判空只能写成 Not 在前：If Not obj Is Nothing Then
    '   VBA 没有 Is Not Nothing 这种写法（C# 才有）

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
    ' UCase$ / Trim$ 名字后的 $ 是 String 类型声明字符：带 $ 的版本返回 String，
    ' 不带 $ 的返回 Variant（更慢，且会把 Null 传下去污染整个表达式），详见 1.2
    Select Case UCase$(Trim$(code))            ' Select 后面只能跟一个表达式
        Case "PASS", "OK"                      ' 一个 Case 可以跟多个值，逗号分隔
            MsgBox "判定为合格"
        Case "FAIL", "NG"
            MsgBox "判定为不合格"
        Case Else                              ' 兜底分支，放最后
            MsgBox "未知结果：[" & code & "]"
    End Select

    ' ★ Case 不是"返回"，只是"候选值"的标签：Select Case 先算一次表达式，
    '   再从上往下跟每个 Case 比，命中就执行那一支的语句，然后跳到 End Select。
    '   VBA 没有 return，想让分支算出点什么带出去，就自己赋值：
    '     Sub 里      → 赋值给外面已经声明好的变量（如下）
    '     Function 里 → 赋值给函数名本身（见 3.1）
    Dim grade As String
    Select Case UCase$(Trim$(code))
        Case "PASS", "OK"
            grade = "合格"
        Case "FAIL", "NG"
            grade = "不合格"
        Case Else
            grade = "未知"
    End Select
    MsgBox "code = [" & code & "] → 判定：" & grade

    ' ---- Case 的三种写法 ----
    Dim n As Long
    n = 7
    Select Case n
        Case 1 To 5                            ' 闭区间
            MsgBox "1~5"
        Case 6, 8                              ' 枚举
            MsgBox "6 或 8"
        Case Is > 10                           ' Is 是占位符，读作"n Is > 10"，
                                               ' 后面只能跟 > >= < <= = <>（见下）
            MsgBox "大于 10"
        Case Else
            MsgBox "其他：" & n
    End Select
    ' ★ 从上往下匹配，命中第一个就走 → 范围有重叠时，窄的（特殊的）写前面
    ' ★ 要同时判断两个变量（如 同时看电压和电流），Select Case 做不到，用 If
    ' ★ 三类写法可以混写在同一个 Case 里，逗号分隔，任一项命中即执行：
    '     Case 1, 3, 5 To 8, Is > 100
    '   注意 To 必须左小右大，写反了（Case 5 To 1）不报错但永不命中
End Sub
