Option Explicit

' ============================================================
' 1.4 数组：Dim 名称(下标) As 类型
'     默认下界为 0（模块顶部写 Option Base 1 则改成 1）
'     ★ 坑：Dim a(3) 是 0~3 共 4 个元素，不是 3 个
'     ★ 和 JS 的区别：VBA 的「二维数组」是一个数组带两个下标（矩形），
'       不是 JS 那种数组套数组；不能写 g(1)(2)，也不能参差
'     ★ 动态数组 = Dim a() + ReDim / ReDim Preserve 控制大小；
'       Preserve 的规则、二维限制、扩容性能单独放在 DemoReDim
'     ★ LBound / UBound 的规则、元素个数算法、维号与边界见 DemoBounds
' ============================================================
Sub DemoArray()
    ' ---- 1.4.1 一维固定数组 ----
    Dim fixedArr(1 To 3) As Integer   ' 明确写死 1~3，共 3 个
    Dim zeroBased(3) As Integer       ' 省略下界 → 0~3，共 4 个（新手最容易数错）

    fixedArr(1) = 10
    fixedArr(2) = 20
    fixedArr(3) = 30
    zeroBased(0) = 100
    zeroBased(3) = 400

    ' ---- 1.4.2 动态数组：用多少开多少 ----
    Dim dynArr() As String            ' 声明时不写长度（固定数组不能 ReDim）
    ReDim dynArr(1 To 2)              ' 运行期再定长度
    dynArr(1) = "Pass"
    dynArr(2) = "Fail"

    ReDim Preserve dynArr(1 To 3)     ' Preserve = 保留已有数据，只扩容
    dynArr(3) = "Retest"              ' 少了 Preserve 的话，前两个元素会被清空
    ' 展开见 DemoReDim：Preserve 只能改最后一维、每次扩容都要整体拷贝、
    ' 二维数组加行会报错 9

    ' ---- 1.4.3 二维数组（行, 列）----
    Dim grid(1 To 2, 1 To 3) As Double ' 2×3 = 6 个格子，必须矩形，不能参差
    grid(1, 1) = 3.3                   ' 第 1 行第 1 列：引脚 1 的第 1 个电压值
    grid(2, 3) = 5#

    MsgBox "fixedArr(2) = " & fixedArr(2) & vbCrLf & _
           "dynArr(3) = " & dynArr(3) & vbCrLf & _
           "grid(1,1) = " & grid(1, 1) & vbCrLf & _
           "一维元素个数：" & (UBound(fixedArr) - LBound(fixedArr) + 1) & vbCrLf & _
           "行数：" & (UBound(grid, 1) - LBound(grid, 1) + 1) & _
           "，列数：" & (UBound(grid, 2) - LBound(grid, 2) + 1)
    ' ★ LBound = 下界，UBound = 上界；元素个数 = UBound - LBound + 1
    ' ★ 第二个参数是"第几维"，不写默认第 1 维：
    '   UBound(grid) = 2（行），UBound(grid, 2) = 3（列）
    ' ★ 别把 UBound 当元素个数：Dim a(3) 的 UBound 是 3，但它是 4 个元素

    ' 二维数组可以一次性刷进工作表，行/列天然对应（ATE 写报告常用）
    ' Range("A1").Resize(2, 3).Value = grid

    ' ---- 1.4.4 简写形式 ----
    ' ① 类型声明字符：Dim 名字 + 符号，数组照样能用
    Dim iArr%(1 To 3)                 ' = Dim iArr(1 To 3) As Integer
    Dim sArr$(1 To 2)                 ' = Dim sArr(1 To 2) As String
    Dim mArr#(1 To 2, 1 To 2)         ' = Dim mArr(1 To 2, 1 To 2) As Double
    Dim dyn$()                        ' 动态 String 数组的简写
    iArr%(1) = 5
    sArr$(1) = "VCC"

    ' ② Array()：一行塞一组值，省去逐个赋值
    Dim codes As Variant
    codes = Array("VCC", "GND", "CLK")
    MsgBox "第 3 个信号：" & codes(2) & _
           "，共 " & (UBound(codes) - LBound(codes) + 1) & " 个"

    ' ③ Split()：字符串直接拆成数组，比手写循环省事
    Dim pins() As String
    pins = Split("VCC,GND,CLK,RESET", ",")
    MsgBox "拆出 " & (UBound(pins) - LBound(pins) + 1) & " 个引脚，第 2 个是 " & pins(1)

    ' ④ Join()：数组拼回字符串（Split 的反操作）
    MsgBox "拼回去：" & Join(pins, " | ")

    ' ⑤ Filter()：按「包含」筛选子数组
    '    Filter(sourcearray, match [, include] [, compare])
    '      sourcearray  一维字符串数组（Split / Array 出来的都行）
    '      match        要找的子串；传 "" 返回整份拷贝
    '      include      True(默认) 保留匹配项；False 保留不匹配项
    '      compare      vbBinaryCompare(0，默认) 区分大小写
    '                   vbTextCompare(1) 不区分大小写
    '    返回：新的 0 基数组，原数组不动
    MsgBox "含 C 的：" & Join(Filter(pins, "C"), ",")                 ' VCC,CLK
    MsgBox "不含 C 的：" & Join(Filter(pins, "C", False), ",")        ' GND,RESET
    MsgBox "忽略大小写：" & Join(Filter(pins, "c", True, vbTextCompare), ",")  ' VCC,CLK

    ' ★ 只做「包含子串」，不是「开头/结尾」，也不是 >3、按长度、正则这类条件
    ' ★ 返回数组永远 0 基（哪怕 pins 是 1 基），要用 LBound/UBound 重新量
    ' ★ 一个都没匹配上时返回空数组：LBound = 0、UBound = -1，
    '   用 UBound(hit) < 0 判断没命中，别直接 hit(0)
    ' ★ compare 省略时按模块的 Option Compare（默认二进制 = 区分大小写）

    Dim hit As Variant
    hit = Filter(pins, "NC")                          ' 故意找一个不存在的
    If UBound(hit) < 0 Then
        MsgBox "没有命中 NC"
    Else
        MsgBox "命中：" & Join(hit, ",")
    End If

    ' Split 的 limit 与 Join 的限制见 5.1

    ' ---- 1.4.5 常用操作 ----
    ' For Each 遍历一维数组（遍历变量必须是 Variant）
    Dim p As Variant, txt As String
    For Each p In codes
        txt = txt & p & " "
    Next
    MsgBox "遍历结果：" & txt

    ' 判断变量是不是数组
    MsgBox "IsArray(codes) = " & IsArray(codes)

    ' Erase：动态数组变回"未分配"（之后 UBound 会报下标越界，必须重新 ReDim 定义数组）
    Erase dynArr
    ReDim dynArr(1 To 1)
    dynArr(1) = "Fresh"
    MsgBox "重新 ReDim 后：" & dynArr(1)

    ' 注意：VBA 不支持数组整体赋值，arr2 = arr1 会编译报错，
    '      要循环逐个拷，或者两边都用 Variant 来装
End Sub


' ============================================================
' 1.4.6 ReDim：动态数组的「重新分配大小」
'   Dim a()                 只声明，不给长度（固定数组不能 ReDim）
'   ReDim a(1 To n)         重新分配；不带 Preserve 会清空已有数据
'   ReDim Preserve a(...)   保留已有元素，只改大小
'   ★ Preserve 不是免费的：每次都要新开一块内存 + 整体拷贝
' ============================================================
Sub DemoReDim()
    ' ---- ① 不带 Preserve：数据静默清空 ----
    Dim sArr() As String
    ReDim sArr(1 To 2)
    sArr(1) = "Pass"
    sArr(2) = "Fail"
    ReDim sArr(1 To 4)                       ' 只改大小，不保留
    MsgBox "不带 Preserve：[" & sArr(1) & "]/[" & sArr(2) & "]"   ' []/[]
    ' ★ 漏写 Preserve 不会报错，只会把数据清掉 —— 排查时很难发现

    ' ---- ② Preserve：旧元素保留，新元素是默认值 ----
    ReDim sArr(1 To 2)
    sArr(1) = "Pass"
    sArr(2) = "Fail"
    ReDim Preserve sArr(1 To 4)              ' 扩容到 4
    MsgBox "扩容后：" & sArr(1) & "/" & sArr(2) & "/[" & sArr(3) & "]"
    ' 输出：Pass/Fail/[]
    ' ★ 新元素的默认值：String = ""、数值 = 0、Variant = Empty、对象 = Nothing
    ' ★ 只在尾部追加，不会插到中间，也不会自动补上一轮的值

    ' ---- ③ 缩容：尾部砍掉，留下的元素不变 ----
    sArr(3) = "Retest"
    sArr(4) = "Skip"
    ReDim Preserve sArr(1 To 2)
    MsgBox "缩容后：" & Join(sArr, ",")       ' Pass,Fail
    ' ★ 被砍掉的数据找不回来；要删中间元素，得先自己把后面的往前挪

    ' ---- ④ 省略下界：ReDim a(5) 是 0~5 共 6 个 ----
    Dim nArr() As Long
    ReDim nArr(5)                             ' Option Base 0 时是 0~5
    MsgBox "ReDim nArr(5) 元素数：" & (UBound(nArr) - LBound(nArr) + 1)   ' 6
    ' ★ 和 Dim a(3) 是同一个坑；结果还取决于模块顶部有没有 Option Base 1
    ' ★ 想精确控制就写 ReDim nArr(1 To 5)，别依赖默认下界

    ' ---- ⑤ Preserve 只能改「最后一维」 ----
    Dim g() As Long
    ReDim g(1 To 2, 1 To 3)                   ' 2 行 3 列
    g(1, 3) = 33
    ReDim Preserve g(1 To 2, 1 To 5)          ' 改第 2 维（最后一维）→ 可以
    MsgBox "加列后 g(1,3)：" & g(1, 3)         ' 33，旧数据保住了
    ' ReDim Preserve g(1 To 4, 1 To 5)        ' ✘ 改第 1 维 → 运行时错误 9
    ' ★ 二维数组想「按行追加」很别扭：Preserve 动不了行（第 1 维）。
    '   常见做法：反过来存 arr(列, 行)、预分配足够行数，或先用 Collection 收集

    ' ---- ⑥ 固定数组不能 ReDim ----
    ' Dim fixedArr(1 To 3) As Long
    ' ReDim fixedArr(1 To 5)                  ' ✘ 编译错误：Array already dimensioned
    ' ★ 要 ReDim，声明时留空括号：Dim a() / Dim a() As Long
    ' ★ 连 Dim 都不写、直接 ReDim arr(1 To 5) 也能编译（VBA 隐式当 Variant 数组），
    '   但类型是 Variant、意图也不清楚，正式代码别省这一行

    ' ---- ⑦ Erase 之后必须重新 ReDim ----
    Erase sArr
    ' MsgBox UBound(sArr)                     ' ✘ 运行时错误 9：下标越界
    ReDim sArr(1 To 1)
    sArr(1) = "Fresh"
    MsgBox "Erase 后重建：" & sArr(1)
    ' ★ Erase 释放内存，动态数组回到「未分配」；UBound/LBound 都会报错 9
    ' ★ 只想清空但保留大小：用 ReDim（不带 Preserve），别用 Erase

    ' ---- ⑧ 循环里逐次 +1 扩容是 O(n²) ----
    ' 反例（数据一多就明显变慢）：
    '   For i = 1 To 10000
    '       ReDim Preserve sArr(1 To i)       ' 每次都新开内存 + 拷贝全部旧元素
    '       sArr(i) = ...
    '   Next
    ' 正确做法 A：两遍扫描 —— 第一遍只数个数，ReDim 一次，第二遍填值
    ' 正确做法 B：按 2 倍扩容，最后再裁到实际长度（下面这段）

    Dim buf() As String, cap As Long, n As Long, i As Long
    cap = 4
    ReDim buf(1 To cap)                       ' 先开一小块
    For i = 1 To 10
        If n = cap Then
            cap = cap * 2
            ReDim Preserve buf(1 To cap)      ' 满了才扩，次数是 log₂
        End If
        n = n + 1
        buf(n) = "Pin" & n
    Next
    If n > 0 And n < cap Then
        ReDim Preserve buf(1 To n)            ' 收尾裁掉多余容量
    End If
    MsgBox "收集到 " & n & " 个：" & Join(buf, ",")
    ' ★ n = 0 时别 ReDim 成 1 To 0，空结果单独约定（返回 Empty / 留 1 个占位）
    ' ★ 数量完全未知又不想扫两遍：先用 Collection 追加，最后 ReDim 一次倒进去

    ' ---- ⑨ 几条容易忽略的规则 ----
    ' ★ ReDim 是可执行语句，只能写在过程里；模块顶部只能 Dim/Private a()
    ' ★ ReDim 不能改类型：Dim a() As String 之后，元素永远是 String
    ' ★ Preserve 只能改上界：改下界会直接报错（文档明确），下界定了就别动
    ' ★ 函数要返回数组：Dim res() As String → ReDim → 填 → 函数名 = res，
    '   函数返回类型写 Variant（见 3.4）
End Sub


' ============================================================
' 1.4.7 LBound / UBound：数组的上下界
'   LBound(a [, 维])  下界（第一个可用下标）
'   UBound(a [, 维])  上界（最后一个可用下标）
'   元素个数 = UBound - LBound + 1        ← 不是 UBound！
'   维号从 1 开始；省略维号 = 第 1 维
' ============================================================
Sub DemoBounds()
    ' ---- ① 一维：下界由声明方式决定 ----
    Dim a(3) As Long                  ' 省略下界 → 0~3（本模块没写 Option Base）
    MsgBox "LBound=" & LBound(a) & "，UBound=" & UBound(a)   ' 0 / 3
    MsgBox "元素个数：" & (UBound(a) - LBound(a) + 1)         ' 4
    ' ★ 元素个数不是 UBound：a 的上界是 3，但它有 4 个元素

    Dim b(1 To 5) As Long             ' 显式写死 → 1~5
    MsgBox "LBound=" & LBound(b) & "，UBound=" & UBound(b)   ' 1 / 5
    ' ★ 循环一律 For i = LBound(a) To UBound(a)：
    '   别写死 0 To 3 / 1 To 5 —— 换成 Option Base 1 或别人传进来的数组都不会错

    ' ---- ② 二维：必须指定第几维 ----
    Dim g(1 To 2, 0 To 3) As Double
    MsgBox "行数：" & (UBound(g, 1) - LBound(g, 1) + 1) & _
           "，列数：" & (UBound(g, 2) - LBound(g, 2) + 1)     ' 2 行 4 列
    ' ★ UBound(g) 只返回第 1 维（行）的上界，不是总元素数
    ' ★ 维号超出实际维数会报运行时错误 9（一维数组写 UBound(a, 2) 就炸）

    ' ---- ③ 不同来源的数组，下界不一样 ----
    Dim s As Variant
    s = Split("VCC,GND,CLK", ",")
    MsgBox "Split 下界：" & LBound(s)          ' 0（Split 永远是 0 基）

    s = Array("VCC", "GND", "CLK")
    MsgBox "Array 下界：" & LBound(s)          ' 0（写 Option Base 1 才变 1）

    ' ★ Range.Value 读「多单元格区域」是二维、1 基：
    '   s = Range("A1:C1").Value → UBound(s, 1) = 1，UBound(s, 2) = 3
    '   （单个单元格读出来是标量，不是数组）

    ' ---- ④ 动态数组：ReDim 之前 / Erase 之后不能问 ----
    Dim d() As Long
    ' MsgBox UBound(d)                         ' ✘ 运行时错误 9：下标越界
    ReDim d(1 To 3)
    MsgBox "ReDim 后 UBound：" & UBound(d)     ' 3
    Erase d
    ' MsgBox UBound(d)                         ' ✘ 又回到未分配，错误 9
    ' ★ UBound/LBound 只能对「已分配」的数组用；
    '   拿到别人传进来的数组，先判 IsArray，再取界

    ' ---- ⑤ 遍历模板：二维必须双层 For ----
    Dim m() As Long, i As Long, j As Long, txt As String
    ReDim m(1 To 2, 1 To 2)
    For i = LBound(m, 1) To UBound(m, 1)
        For j = LBound(m, 2) To UBound(m, 2)
            m(i, j) = i * 10 + j
            txt = txt & m(i, j) & " "
        Next
    Next
    MsgBox "二维遍历：" & txt                   ' 11 12 21 22
    ' ★ For Each 也能遍历二维数组的全部元素，但拿不到行/列号；
    '   要行列号就老老实实双层 For
    ' ★ LBound/UBound 返回 Long，别声明成 Integer（元素多时会溢出）
End Sub


' ============================================================
' 1.4.8 数组常用操作：VBA 没有 push / sort / slice，得自己写
'   JS 对照：
'     arr.length          UBound(arr) - LBound(arr) + 1
'     arr.push(x)         ReDim Preserve arr(UBound + 1)（或预分配 + 计数）
'     arr.pop()           取 arr(UBound)，再 ReDim Preserve 缩 1
'     arr.includes(x)     循环比较
'     arr.indexOf(x)      循环，找不到返回 -1（VBA 没有内置）
'     arr.slice(a, b)     循环拷到新数组
'     arr.concat(b)       循环拷两段
'     arr.reverse()       首尾对调
'     arr.sort()          冒泡；数据大用 WorksheetFunction.Sort
'     arr.join(",")       Join
'     arr.filter(fn)      Filter 只做「包含子串」，条件过滤要自己循环
'     arr.map(fn)         循环
'     Array.isArray(x)    IsArray
'     [...new Set(arr)]   用 Dictionary 去重（见 5.3）
'     delete arr[i]       没有；只能挪元素 + 缩容
' ============================================================
Sub DemoArrayOps()
    ' ---- ① 追加 / 弹出：只能动尾部 ----
    Dim pins() As String
    ReDim pins(1 To 1)
    pins(1) = "VCC"

    ReDim Preserve pins(1 To UBound(pins) + 1)
    pins(UBound(pins)) = "GND"                   ' push
    ReDim Preserve pins(1 To UBound(pins) + 1)
    pins(UBound(pins)) = "CLK"                   ' push
    MsgBox "追加后：" & Join(pins, ",")          ' VCC,GND,CLK

    Dim lastPin As String, i As Long, j As Long
    lastPin = pins(UBound(pins))                 ' pop：先取最后一个
    If UBound(pins) > LBound(pins) Then
        ReDim Preserve pins(LBound(pins) To UBound(pins) - 1)
        MsgBox "弹出 " & lastPin & "，剩：" & Join(pins, ",")   ' CLK，剩 VCC,GND
    Else
        Erase pins
        MsgBox "弹出 " & lastPin & "，数组已空"
        Exit Sub                                 ' 后面都依赖 pins，空了就别往下跑
    End If
    ' ★ 删中间元素没有内置方法：把后面的往前挪，再 ReDim Preserve 缩 1

    ' ---- ② 查找：没有 indexOf / includes ----
    Dim target As String, idx As Long
    target = "GND"
    idx = -1                                     ' 和 JS 一样用 -1 表示找不到
    For i = LBound(pins) To UBound(pins)
        If pins(i) = target Then
            idx = i
            Exit For
        End If
    Next
    MsgBox "indexOf(GND) = " & idx               ' 2
    ' ★ -1 是「我们自己约定」的，VBA 没有这个函数
    ' ★ 字符串默认区分大小写；忽略大小写用 StrComp(pins(i), target, vbTextCompare) = 0

    ' ---- ③ 切片：没有 slice，循环拷 ----
    Dim subArr() As String
    ReDim subArr(1 To 2)
    For i = 1 To 2
        subArr(i) = pins(i)
    Next
    MsgBox "前 2 个：" & Join(subArr, ",")        ' VCC,GND
    ' ★ 要通用就写成 Function ArraySlice(...) As Variant（见 3.4 返回数组）

    ' ---- ④ 合并：没有 concat，循环拷两段 ----
    Dim more As Variant, merged() As String, k As Long, nTotal As Long
    more = Split("RESET,NC", ",")
    nTotal = (UBound(pins) - LBound(pins) + 1) + (UBound(more) - LBound(more) + 1)
    ReDim merged(1 To nTotal)
    For i = LBound(pins) To UBound(pins)
        k = k + 1
        merged(k) = pins(i)
    Next
    For i = LBound(more) To UBound(more)
        k = k + 1
        merged(k) = CStr(more(i))
    Next
    MsgBox "合并：" & Join(merged, ",")          ' VCC,GND,RESET,NC

    ' ---- ⑤ 反转：首尾对调，原地改 ----
    Dim tmp As String
    For i = 1 To UBound(merged) \ 2
        j = UBound(merged) - i + 1
        tmp = merged(i)
        merged(i) = merged(j)
        merged(j) = tmp
    Next
    MsgBox "反转：" & Join(merged, ",")          ' NC,RESET,GND,VCC
    ' ★ 这里默认数组从 1 开始；0 基要自己换算，别写死 1

    ' ---- ⑥ 排序：没有 sort，先用冒泡 ----
    Dim volt As Variant, t As Variant
    volt = Array(3.3, 1.25, 5, 0)
    For i = LBound(volt) To UBound(volt) - 1
        For j = i + 1 To UBound(volt)
            If volt(j) < volt(i) Then
                t = volt(i)
                volt(i) = volt(j)
                volt(j) = t
            End If
        Next
    Next
    MsgBox "升序：" & Join(volt, ", ")           ' 0, 1.25, 3.3, 5
    ' ★ 数据量大别用冒泡（O(n²)）：用 WorksheetFunction.Sort / Small，
    '   或者写进工作表排完再读回（4.3）
    ' ★ 字符串排序把 < 换成 StrComp(...) < 0；混合类型会报「类型不匹配」

    ' ---- ⑦ 条件过滤：Filter 只认子串，不认条件 ----
    Dim vals As Variant, big() As Variant, cnt As Long, m As Long
    vals = Array(3.31, 0.02, 5.5, 1.79)
    For i = LBound(vals) To UBound(vals)
        If vals(i) > 3 Then cnt = cnt + 1
    Next
    If cnt > 0 Then
        ReDim big(1 To cnt)                      ' 两遍扫描：先数个数，再 ReDim 一次
        For i = LBound(vals) To UBound(vals)
            If vals(i) > 3 Then
                m = m + 1
                big(m) = vals(i)
            End If
        Next
        MsgBox ">3 的值：" & Join(big, ", ")     ' 3.31, 5.5
    Else
        MsgBox "没有 >3 的值"
    End If
    ' ★ Filter(arr, "VCC") 是子串匹配，不是条件；>3、长度、正则都得自己循环

    ' ---- ⑧ 一维 ↔ 二维：Transpose（仅 Excel 宿主）----
    Dim oneD As Variant, twoD As Variant
    oneD = Array(10, 20, 30)
    twoD = Application.Transpose(oneD)           ' 1D → 3×1 二维
    MsgBox "转成 " & UBound(twoD, 1) & "×" & UBound(twoD, 2)   ' 3×1
    ' ★ 写一列时用：Range("A1").Resize(3, 1).Value = twoD
    ' ★ 反过来：单列 Range.Value 是 n×1 二维，Transpose 一次可转 1D，
    '   之后 Join / Filter 才能用
    ' ★ Transpose 最多 65536 个元素，超了会失败；几十万行别用它
End Sub
