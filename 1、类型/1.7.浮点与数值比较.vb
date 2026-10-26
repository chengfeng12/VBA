Option Explicit

' ============================================================
' 1.7 浮点与数值比较：ATE 判限最容易翻车的地方
'     核心结论：两个 Double 永远不要用 = 比较
'     判限的正确姿势：Abs(实测 - 规格) <= 容差，而不是 实测 = 规格
' ------------------------------------------------------------
' 和 JS 一模一样（都是 IEEE 754 双精度）：
'     0.1 + 0.2 === 0.3        // false
'     0.1 + 0.2 <> 0.3         ' True
' 所以套路也通用：Math.abs(a - b) < eps  →  Abs(a - b) < eps
' ============================================================
Sub DemoFloat()

    ' ---- 现象：看着相等的两个数其实不等 ----
    Dim a As Double, b As Double
    a = 0.1 + 0.2
    b = 0.3
    MsgBox "0.1+0.2 = " & a                              ' 0.3（显示被四舍五入过）
    MsgBox "a = b 吗？" & (a = b)                        ' False
    MsgBox "差值：" & (a - b)                             ' 5.55111512312578E-17
    ' ★ MsgBox 只显示 15 位左右，看着是 0.3 而已，内部并不是

    ' ---- 正确比较：容差（eps）----
    MsgBox "容差比较：" & (Abs(a - b) < 0.0000001)       ' True
    ' 容差按量程定：判 3.3 V 用 1e-6 就够了，
    ' 判 0.001 A 这种小量程要更小，否则 1e-6 会把整段规格吃掉

    ' ---- 判定"是不是 0" ----
    Dim x As Double
    x = 1E-12
    MsgBox "x = 0 吗？" & (x = 0)                        ' False
    MsgBox "按容差判零：" & (Abs(x) < 1E-9)              ' True

    ' ---- 判限：上下限用 >= / <=，别用 = ----
    Dim measured As Double, lo As Double, hi As Double
    measured = 3.2999999999999
    lo = 3.3
    hi = 3.6
    MsgBox "直接比较：" & (measured >= lo)                ' False ← 假 Fail
    MsgBox "带容差：" & (measured >= lo - 0.000001)       ' True  ← 正确
    ' ★ ATE 里"边界值刚好卡在 limit 上"的场景必有，
    '   不加容差就会出现"同一个料，换台机台判 Fail"的经典扯皮
End Sub


' ============================================================
' 舍入：VBA 的 Round 是"银行家舍入"，不是四舍五入
'   Round(2.5) = 2      Round(3.5) = 4      ← 逢 .5 取偶数
'   CInt / CLng 同样走银行家舍入
'   要"四舍五入"就用 WorksheetFunction.Round（Excel 的 ROUND 函数）
' ============================================================
Sub DemoRound()
    MsgBox "Round(2.5) = " & Round(2.5)                  ' 2（不是 3）
    MsgBox "Round(3.5) = " & Round(3.5)                  ' 4
    MsgBox "CInt(2.5)  = " & CInt(2.5)                   ' 2
    MsgBox "CInt(3.5)  = " & CInt(3.5)                   ' 4
    ' Excel 的 ROUND 才是"四舍五入"：
    MsgBox "Excel ROUND(2.5) = " & Application.WorksheetFunction.Round(2.5, 0)   ' 3
    ' ★ 报表数字对不上时，先怀疑这里，再怀疑浮点

    ' ---- Int 与 Fix 对负数的方向不同 ----
    MsgBox "Int(-3.5) = " & Int(-3.5)                    ' -4（向下，朝负无穷）
    MsgBox "Fix(-3.5) = " & Fix(-3.5)                    ' -3（向零截断）

    ' ---- 格式化输出：Format 会按系统区域设置加千分位/换小数点 ----
    MsgBox "Format 两位：" & Format$(3.14159, "0.00")     ' 3.14
    ' ★ 小数点不是英文句点的系统上（区域设置为德语/法语等），
    '   Format 会输出 "3,14"，直接拼进 CSV 会多出一列 ——
    '   写文件前统一 Replace(s, ",", ".") 或改用 CStr + 手动截位
End Sub


' ============================================================
' 精度常识：Double 约 15~16 位有效数字
'   超过这个位数，低位就是噪声；累加几万次后误差会累积
' ============================================================
Sub DemoPrecision()
    ' 大数吃小数：1e16 + 1 = 1e16（1 被吞掉）
    MsgBox CDbl(1E+16) + 1 = 1E+16                       ' True

    ' 累加误差：0.1 加 10 次不等于 1
    Dim sum As Double, i As Long
    For i = 1 To 10
        sum = sum + 0.1
    Next
    MsgBox "累加 10 次 0.1 = " & sum                      ' 0.999999999999999 或 1
    MsgBox "等于 1 吗？" & (sum = 1)                       ' False
    ' 对策：整数化（乘 1000 用 Long 累加）或最后再舍入，别边加边判

    ' Currency 是 4 位小数的定点数，金额类用它不会漂
    Dim money As Currency
    money = 0.1@ + 0.2@
    MsgBox "Currency：0.1 + 0.2 = " & money & "，等于 0.3 吗？" & (money = 0.3@)   ' True
End Sub
