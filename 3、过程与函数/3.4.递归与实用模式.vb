Option Explicit

' ============================================================
' 3.4 递归与实用模式
'     递归、Static 记忆、返回数组、命名冲突
' ============================================================
Sub DemoRecursion()
    ' ---- 递归：函数自己调自己，必须有"终止条件"，否则栈溢出 ----
    MsgBox "5! = " & Fact(5)
    ' ★ Long 最大 21 亿，只能装到 12!（12! = 479,001,600），
    '   13! 就溢出了；再大就得换成 Double（会丢精度）或 Currency
    ' MsgBox "13! = " & Fact(13)          ' ← 取消注释会报"溢出"

    ' ---- Static：函数内的静态变量，跨调用保留 ----
    MsgBox "第 1 次：" & NextSeq() & "，第 2 次：" & NextSeq() & "，第 3 次：" & NextSeq()
    ' ★ Static 适合"计数器 / 缓存"；但值会一直留着，
    '   调试时发现数值"莫名其妙不平"，先想这里

    ' ---- 返回数组：只能靠 Variant 装 ----
    Dim pins As Variant
    pins = ReadPins()
    MsgBox "共 " & (UBound(pins) - LBound(pins) + 1) & " 个引脚，第 1 个是 " & pins(0)
    ' ★ Function Foo() As String() 这种写法是"非法的"，
    '   要返回数组就声明成 As Variant，再用 Array() 塞进去

    ' ---- 命名冲突：VBA 没有命名空间 ----
    ' 两个模块都写了 ToDouble → 编译期报"二义性名称"，两个都不能用
    ' 规避：过程名带模块前缀（UtilToDouble / ReportToDouble），
    '       别用 Range / Date / Left 这类内置名当自己的过程名
End Sub


' ============================================================
' 阶乘：n <= 1 是终止条件，否则 n * (n-1)!
' 递归一定要先写终止条件，再写递归调用
' ============================================================
Function Fact(ByVal n As Long) As Long
    If n <= 1 Then
        Fact = 1
        Exit Function
    End If
    Fact = n * Fact(n - 1)
End Function


' ============================================================
' Static 局部变量：只初始化一次，之后一直保留
' 等价于模块级变量，但作用范围只在本函数内
' ============================================================
Function NextSeq() As Long
    Static cnt As Long
    cnt = cnt + 1
    NextSeq = cnt
End Function


' ============================================================
' 返回数组：声明成 Variant，用 Array() 造数组
' ============================================================
Function ReadPins() As Variant
    ReadPins = Array("VCC", "GND", "CLK")
End Function
