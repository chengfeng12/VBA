Option Explicit

' ============================================================
' 1.6 工具函数：任意输入 → Double（规避 CDbl 的区域设置坑）
' ------------------------------------------------------------
' 【为什么别的模块能直接用 ToDouble，不用 import？】
'   VBA 没有模块系统，也没有 import 语句。
'   一个 .xlsm 文件 = 一个 VBAProject，工程里的所有模块（标准模块、
'   类模块、窗体、ThisWorkbook、Sheet1）共享同一个全局命名空间，
'   编译期一次性把所有名字解析完。
'
'   所以：本文件粘进工程 = 已经"注册"好了，任何模块直接喊名字就能用。
'   反过来：漏粘本模块 → 编译期立刻报「Sub 或 Function 未定义」，
'          这才是 JS 里"忘记 import"的等价场景（只是它发生在编译期）。
'
'   作用域规则：
'     Sub / Function 不写修饰符  = Public，全工程可见（默认值）
'     前面加 Private             = 仅本模块可见
'     模块首行加 Option Private Module = 整个模块对外部工程不可见
'   想只给本模块用，就把下面改成 Private Function ToDouble(...)
'
'   真正需要"引用"的只有跨工程：工具 → 引用 勾对象库，
'   或者用 Application.Run "别的文件.xlsm!宏名" 调用。
' ------------------------------------------------------------
' 【函数行为】
'   参数故意用 Variant 收，字符串 / 数值 / 单元格 Value 都能直接丢进来
'   1. 字符串：去首尾空格、去千分位逗号，再用 Val 解析
'   2. 数值  ：直接 CDbl，不走字符串解析
'   3. 其他（Empty / Null / Date / Object）：抛错，绝不静默返回 0
'   注意：IsNumeric(日期) 在 VBA 里返回 True，所以这里显式排除 vbDate
' ============================================================
Function ToDouble(ByVal v As Variant) As Double
    Dim clean As String

    Select Case VarType(v)
        Case vbString                                  ' "22.5" / "1,234.5"
            clean = Trim$(v)
            clean = Replace(clean, ",", "")            ' 去千分位
            If Not IsNumeric(clean) Then GoTo Bad
            ToDouble = Val(clean)                      ' 刻意不用 CDbl（本机会溢出）

        Case vbByte, vbInteger, vbLong, vbSingle, _
             vbDouble, vbCurrency, vbDecimal            ' 本来就是数值
            ToDouble = CDbl(v)

        Case Else                                       ' Empty / Null / Date / Object
            GoTo Bad
    End Select
    Exit Function

Bad:
    Err.Raise vbObjectError + 1, "ToDouble", "不是合法数字：[" & CStr(v) & "]"
End Function


' 反方向：数值 → 固定 n 位小数的字符串（避免科学计数法，ATE 报告里常用）
Function ToFixed(ByVal d As Double, Optional ByVal digits As Long = 4) As String
    ToFixed = Format$(d, "0." & String$(digits, "0"))
End Function
