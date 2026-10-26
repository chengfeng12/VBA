Option Explicit

' ============================================================
' 3.1 Sub 与 Function
'     Sub      做一件事，没有返回值
'     Function 算出一个值，用"给函数名赋值"的方式返回
'     两者不写修饰符就是 Public（全工程可见），加 Private 只给本模块用
' ============================================================
Sub DemoProcBasics()

    ' ---- 调 Sub：两种正确写法 ----
    SayHello                                  ' 无参数：只写名字，不写括号
    Call SayPin("VCC", 3.3)                   ' 有参数：Call + 括号，参数再多也不出错
    SayPin "GND", 0#                          ' 有参数：不带 Call，就不能加括号

    ' ★ 坑：SayPin("GND", 0) 这种"不带 Call 又加括号"的写法会编译报错
    '   （VBA 会以为你想取返回值）。多参数调用统一写 Call 最省心

    ' ---- 调 Function：当值用，必须加括号 ----
    MsgBox "3.3 V 判 " & IIf(JudgeVolt(3.3, 3#, 3.6), "Pass", "Fail")

    Dim total As Double
    total = SumVolt(3.3, 3.31, 3.29)
    MsgBox "三个采样点之和：" & total

    ' ---- 用循环 + 提前返回找"第一个不合格的点"（和第 2 章接上）----
    MsgBox "第一个超上限的位置：" & FirstOverLimit(Array(3.3, 3.31, 3.8), 3.6)
    ' 返回 3（第 3 个采样点 3.8 超了 3.6）；没找到时返回 -1

    ' ---- 返回值的坑：忘了给函数名赋值 ----
    ' VBA 不会报错，直接返回该类型的默认值：
    '   Boolean → False，数值 → 0，String → ""，对象 → Nothing
    ' 这是最常见的静默 bug：条件写错了不报错，只是永远返回 False
    '
    ' ★ 函数名在过程内就是一个普通变量，所以递归里能左右同时出现：
    '     Fact = n * Fact(n - 1)      ' 右边读旧值、左边写新值（见 3.4）
    '   也正因为它是变量，"赋值"不等于"返回"，中间还能继续改

    ' ---- 返回对象：必须用 Set，不能直接等号 ----
    '   正确：Set GetSelf = ThisWorkbook
    '   错误：GetSelf = ThisWorkbook     ' 想赋对象却走了默认属性，结果常常是 Nothing 或报错
    '   JS 里 return obj 一行搞定，VBA 要记住"对象赋值一律 Set"（4.1 会再讲）
    MsgBox "Set 返回对象：" & GetSelf().Name

    ' ---- Function 也能当 Sub 用：单独一行调用，返回值直接丢掉 ----
    JudgeVolt 3.3, 3#, 3.6                    ' 和 JS 里 fn() 不接收返回值一个意思
    ' 但反过来不行：Sub 不能出现在表达式里（它没有值）

    ' ---- Public / Private ----
    ' 不写修饰符 = Public，别的模块直接按名字调用（见 1.6 关于"没有 import"的说明）
    ' 前加 Private = 只有本模块能调，适合纯内部工具
    MsgBox "内部工具：" & InternalTag()

    ' ---- Static 修饰过程：里面所有局部变量自动变成 Static ----
    '   写 Static Sub / Static Function，等于给过程内每个 Dim 都加了 Static，
    '   省去逐个声明；代价是状态更难追踪，一般只给"工具函数"用
    ' ---- 还有 Property 过程（Get / Let / Set）----
    '   类模块 / 窗体模块里可以用 Property 把字段包成属性，
    '   但标准模块里不允许写 Property，日常脚本用不到，先把 Function 用熟
End Sub


' ============================================================
' 普通 Sub：只干活，不返回值
' ============================================================
Sub SayHello()
    MsgBox "你好，过程。"
End Sub


Sub SayPin(ByVal pin As String, ByVal volt As Double)
    MsgBox "引脚 " & pin & " = " & volt & " V"
End Sub


' ============================================================
' Function：返回值写在"函数名"上，不是 return
' ============================================================
Function JudgeVolt(ByVal v As Double, ByVal lo As Double, ByVal hi As Double) As Boolean
    JudgeVolt = (v >= lo And v <= hi)          ' 最后一个表达式就是返回值
End Function


Function SumVolt(ByVal a As Double, ByVal b As Double, ByVal c As Double) As Double
    SumVolt = a + b + c
End Function


' 循环里找到就 Exit Function 提前返回，找不到时显式返回 -1
' （不写的话默认返回 0，调用方会误以为"第 0 个"是结果）
Function FirstOverLimit(ByVal arr As Variant, ByVal hi As Double) As Long
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) > hi Then
            FirstOverLimit = i
            Exit Function
        End If
    Next
    FirstOverLimit = -1
End Function


' 返回对象：Set 函数名 = 对象，少写 Set 就会踩坑
Function GetSelf() As Object
    Set GetSelf = ThisWorkbook
End Function


' Private：只在本模块可见，别的模块喊不到这个名字
Private Function InternalTag() As String
    InternalTag = "私有函数（Private），出了本模块就不存在"
End Function
