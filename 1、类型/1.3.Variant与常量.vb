Option Explicit

' ============================================================
' 1.3 Variant、常量、类型转换
'     Variant = 万能类型（不写 As 就是 Variant）
'     Const   = 常量，声明后不能改
'     Cxxx()  = 类型转换函数
' ============================================================
Sub DemoVariantAndConst()
    Dim v                     ' 不写 As，默认就是 Variant，什么都能装
    v = 123
    MsgBox "现在是数字：" & (v + 1) & "，VarType = " & VarType(v) & "（2 = vbInteger）"
    v = "文本"                ' 同一变量换个类型照样用
    MsgBox "现在是文本：" & v & "，长度 " & Len(v) & "，VarType = " & VarType(v) & "（8 = vbString）"
    v = Array("VCC", "GND")    ' 还能装数组
    MsgBox "现在是数组，VarType = " & VarType(v) & "（8204 = vbArray + vbVariant）"

    ' 注意：Dim a, b As Integer 里只有 b 是 Integer，a 是 Variant
    '      要都声明就得写成 Dim a As Integer, b As Integer

    ' Const 常量：声明后不能改，适合放阈值、路径、测试参数
    Const MAX_RETRY As Long = 3
    Const VOLT_TOL As Double = 0.05
    ' MAX_RETRY = 5   ' ← 取消注释会报错：常量不能重新赋值

    ' ---- 常用类型转换 ----
    Dim t As String
    t = "22.5"

    ' 整数类：会取整，.5 时按银行家舍入取偶（CInt(22.5) = 22，不是 23）
    MsgBox "CInt: " & CInt(t)
    MsgBox "CLng: " & CLng(t)

    ' 浮点类：CDbl 是按"系统区域设置"去认小数点的
    '   本机实测 CDbl("22.5") 会报 运行时错误 6（溢出），连手写常量都炸，
    '   属于环境问题（区域设置里的"小数点符号"不是英文句点 .），不是代码问题
    ' MsgBox "CDbl: " & CDbl(t)      ' ← 这行在本机会炸，先注释掉

    ' 替代方案 1：Val —— 固定按英文 "." 解析，不看区域设置，永不抛异常
    MsgBox "Val: " & Val(t)

    ' 替代方案 2（推荐）：先校验再转换，脏数据不会静默变成 0
    '   ToDouble 见 1.6.工具函数.vb
    '   ★ 这里没有任何 import / 引用动作，直接按名字调用就行，原因：
    '     VBA 一个 .xlsm = 一个工程（VBAProject），工程内所有模块共享同一个
    '     命名空间；Sub/Function 不写修饰符就是 Public（全工程可见），
    '     编译期一次性解析所有名字，不像 JS 那样需要显式导入。
    '     唯一会报错的情况是"1.6 这个模块没粘进工程" → 编译期报
    '     「Sub 或 Function 未定义」，那才是 JS 里"忘记 import"的等价场景。
    '     想反过来限制成只给本模块用，就在 1.6 里把 Function 改成 Private Function。
    '     另外 VBA 没有包管理，重名即冲突：两个模块都有 ToDouble 会报"二义性名称"。
    If IsNumeric(t) Then
        MsgBox "校验后转换: " & ToDouble(t)
    Else
        MsgBox "不是合法数字：[" & t & "]"
    End If

    MsgBox "CStr: " & CStr(123)
    MsgBox "CBool: " & CBool(1) & "（非 0 即 True）"
    MsgBox "IsNumeric(""abc""): " & IsNumeric("abc")

    ' 一句话记法：
    '   字符 → 整数  用 CInt / CLng（会取整）
    '   字符 → 浮点  别用 CDbl，用 Val 或 CDec（本机 CDbl 有坑）
    '   取整数部分  Fix（向 0 取整）/ Int（向下取整），负数时行为不同

    ' 判断变量到底是不是空的
    Dim x
    MsgBox "IsEmpty(x): " & IsEmpty(x)
End Sub
