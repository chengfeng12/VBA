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

    ' ---- 和 JS 的 for 对照 ----
    '   JS ：for (let i = 1; i <= 3; i++)   ← 初始化 / 条件 / 步进 三段全手写
    '   VBA：For i = 1 To 3 [Step 1]         ← 只写初值终值，步进内置、比较方向自动
    ' （Step 为正用 <=，为负自动用 >=，不用自己判方向）
    ' 三个关键差异：
    '   ① 终值只求值一次：循环里改 n 不会改变循环次数（JS 每轮都重新比对条件）
    '   ② i 是过程级变量，循环结束后仍然存在（JS 的 let i 出循环即销毁）
    '   ③ 没有 continue，也没有 ++；要"跳过本轮"只能把语句包进 If，
    '      要一次跳出多层循环只能用标志位或抽成函数（见 3.1）
    Dim n3 As Long
    n3 = 3
    For i = 1 To n3
        n3 = 10                        ' 改了也没用，终值在进入循环时就定了
    Next
    MsgBox "循环里把终值改成 10，实际仍只跑了 3 次，i = " & i

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

    ' ---- 两端的值不限于字面量 ----
    ' "1 To 10" 不是"区间类型"，只是两个求值槽，可以写变量 / 函数调用 / 运算：
    '     For i = LBound(a) To UBound(a)      遍历数组的标准写法
    '     For i = start To start + 5          端点是变量
    ' ★ 只能写一段：For i = 1 To 5, 8 To 10 是非法的（逗号多段是 Select Case 的语法）
    '   要跳段只能 Step 等差，或在循环里用 If 挑
    ' ★ 方向由 Step 决定，不看两端大小：For i = 10 To 1（Step 默认 +1）一次都不执行，
    '   不报错、静默跳过 —— 顺序写反时表现就是"代码好像没跑"

    ' ---- 嵌套循环：外层行、内层列，遍历二维数组（1.4 的 grid）----
    Dim grid(1 To 2, 1 To 3) As Double
    Dim r As Long, c As Long
    grid(1, 1) = 3.3:  grid(1, 2) = 3.31: grid(1, 3) = 3.29
    grid(2, 1) = 0#:   grid(2, 2) = 0.01: grid(2, 3) = 0#

    For r = 1 To 2
        For c = 1 To 3
            grid(r, c) = grid(r, c) * 1000        ' 单位 V → mV
        Next c                                     ' Next 后面写变量名，嵌套时不会错配
                                                   ' （可省略，单写 Next 会自动配最近那个 For，
                                                   '   但这样 VBA 就不校验配对，层数写错也查不出来）
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
    ' ---- 和 JS 对照 ----
    '   VBA: Exit For              ≈  JS: break
    '   VBA: （没有）               ≈  JS: continue   ← 想跳过本轮只能把语句包进 If
    '   JS 的 labeled break（break outer;）VBA 也没有，
    '   多层一起退只能靠标志位，或把内层循环抽成函数用 Exit Function 返回（3.1）
    '   注：VB.NET 有 Continue For，VBA 没有，别混用

    ' ---- For Each：不关心下标时更省事 ----
    Dim codes As Variant, p As Variant, txt As String
    codes = Array("VCC", "Gnd", "CLK")
    For Each p In codes
        txt = txt & p & " "
    Next
    MsgBox "For Each → " & txt
    ' ★ 坑：For Each 的循环变量必须是 Variant（或对象类型），
    '   写成 Dim p As String 会报"需要对象"，因为 VBA 不知道数组元素的真实类型
    ' ★ For Each 只适合"只读遍历"：要改元素值还得用下标 For
    ' ★ 集合（Worksheets / ListObjects）也能 For Each，变量写成对应对象类型，见第 4 章
    ' ---- 和 JS 对照 ----
    '   VBA: For Each p In codes   ≈  JS: for (const p of codes)
    '   ① 循环变量必须是 Variant / 对象类型（JS 里 const 就行，没这限制）
    '   ② 只能只读：改 p 不会写回数组（JS 的 for...of 也一样，要改还得用下标）
    '   ③ 拿不到下标：要序号得自己加计数器（JS 可用 entries() 直接拿 index）
    '   ④ 没有 map / filter / forEach 链式调用；过滤用 Filter()（见 1.4），
    '      映射（逐个转换）只能手写循环
    '   ⑤ 遍历集合（Worksheets 等）时下标从 1 开始，JS 数组从 0 开始

    ' ---- 同一件事的另一种写法：下标 For ----
    ' Array() 生成的数组下界是 0，所以 codes 可用下标是 0~2，codes(3) 会越界。
    ' 但不要自己假设下界：Dim a(n) 的下界会随模块顶部的 Option Base 变（0 或 1），
    ' 一律用 LBound / UBound 取边界，元素个数 = UBound - LBound + 1。
    ' 注意：LBound/UBound 是数组专用，集合（Worksheets）没有，只有 .Count。
    ' 选择标准：
    '   只读遍历            → For Each
    '   要改元素 / 要序号   → 下标 For
    '   遍历对象集合        → For Each（集合没有 LBound/UBound，只有 .Count，且从 1 开始）
    Dim k As Long
    For k = LBound(codes) To UBound(codes)
        codes(k) = UCase$(codes(k))           ' 这次是真的写回数组
    Next k
    MsgBox "下标遍历并改写后：" & Join(codes, " / ")
End Sub
