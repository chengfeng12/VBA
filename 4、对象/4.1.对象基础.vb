Option Explicit

' ============================================================
' 4.1 对象基础：Set / Nothing / Is / With / 默认属性
'     和 JS 的区别：JS 的对象赋值 obj2 = obj1 也是引用，
'     但 VBA 赋值对象时必须写 Set，否则直接编译报错
' ============================================================
Sub DemoObjectBasics()
    Dim ws As Worksheet
    Dim rng As Range

    ' ---- 声明与赋值：对象必须用 Set ----
    Set ws = ThisWorkbook.Worksheets(1)      ' 代码所在工作簿的第 1 张表
    Set rng = ws.Range("A1")

    ' 错误写法：ws = ThisWorkbook.Worksheets(1)   → 编译报错
    ' 记忆法：Set 是"把变量指向那个对象"，不是"把对象内容拷给变量"

    MsgBox "当前引用：" & ws.Name & " 的 " & rng.Address

    ' ---- 对象变量之间是"引用"，不是拷贝 ----
    Dim ws2 As Worksheet
    Set ws2 = ws                    ' ws2 和 ws 指向同一张表
    ' ws2.Name = "改个名"           ' ← 取消注释后 ws.Name 会一起变（本来就是同一张表）

    ' ---- Nothing 与 Is ----
    Dim tmp As Worksheet
    MsgBox "未赋值时为 Nothing：" & (tmp Is Nothing)     ' True
    Set tmp = ws
    If Not tmp Is Nothing Then MsgBox "已指向：" & tmp.Name

    ' 注意：判断对象必须用 Is（tmp Is Nothing），不能写 tmp = Nothing
    '       Nothing 是"不指向任何对象"，不是空字符串也不是 0

    ' ---- 默认属性：Range("A1") = 1 其实等于 Range("A1").Value = 1 ----
    ' ws.Range("A1") = 3.3          ' ← 真写入，默认属性就是 Value
    ' MsgBox ws.Range("A1")         ' ← 读取时默认属性也是 Value
    ' 建议永远显式写 .Value，省得别人（和三个月后的自己）猜
    ' ★ 默认属性在比较里更阴：If ws.Range("A1") = "" 读的是 .Value，
    '   单元格里是错误值 #N/A 时整个 If 直接抛"类型不匹配"，先判 IsError

    ' ---- 判断对象类型：TypeOf ... Is（≈ JS 的 instanceof）----
    ' If TypeOf ws Is Worksheet Then MsgBox "确实是工作表"
    ' MsgBox TypeName(ws)                      ' "Worksheet"，调试时比 Is 直观
    ' ★ 只对对象有效；对 Nothing 用 TypeOf 返回 False，不会报错

    ' ---- ★ Dim ... As New 的陷阱：Is Nothing 从此永远失效 ----
    ' Dim col As New Collection      ' 这种声明会在"第一次用到"时自动 New 一个
    ' MsgBox (col Is Nothing)        ' 永远 False —— 想靠 Is Nothing 判"还没建"会失效
    ' 正确写法：Dim col As Collection : Set col = New Collection（需要时再 Set）

    ' ---- With：省掉重复的对象前缀 ----
    With ws.Range("B2")
        '.Value = "VCC"             ' 真写入，先注释
        MsgBox "地址：" & .Address & vbCrLf & _
               "行：" & .Row & "，列：" & .Column
    End With
    ' With 块里以 . 开头的是 With 后面那个对象；不写 . 就是普通变量/过程
    ' ★ 和 JS 的 with 不是一回事：JS 那个是把对象属性塞进作用域（严格模式禁用），
    '   VBA 的 With 只是"省掉对象前缀"，块里照样能用全局变量和过程名

    ' ---- 嵌套：内层的 . 指向最近的那一层 ----
    With ws.Range("B2")
        With .Font                                  ' .Font 属于 Range
            .Bold = True                            ' 属于 Font
            .Size = 10
        End With
        ' 出了内层，. 又回到 Range
        MsgBox "外层还能用：" & .Address
    End With

    ' ---- 最常见的静默 bug：漏写一个点 ----
    ' With ws
    '     Cells(1, 1).Value = 1        ' ✗ 漏点 → 操作的是【当前活动表】，不是 ws
    '     .Cells(1, 1).Value = 2       ' ✓ 带点才是 ws 上的单元格
    ' End With
    ' ★ 这类错误不报错，只是写错了表 —— 双表操作时特别难发现

    ' ---- 什么时候值得用 ----
    '   连续访问同一对象的 3 个以上成员时用，读起来清爽（尤其设置格式那种长链）；
    '   对 Excel COM 还有实打实的性能收益：每写一次前缀都要重新解析对象，
    '   With 只解析一次
    ' ---- 什么时候别用 ----
    '   块里只有一两行时没必要；块太长（几十行）反而看不出当前操作的是谁

    ' ---- 链式引用：能用但难维护 ----
    ' Application.Workbooks("book.xlsm").Worksheets("Sheet1").Range("A1").Value
    ' 中间任何一环出错都定位不到具体是哪一级，建议逐级 Set 成变量，好调试

    ' ---- 释放引用 ----
    Set tmp = Nothing
    Set ws2 = Nothing
    Set rng = Nothing
    Set ws = Nothing
    ' 小脚本里不写也能跑（过程结束自动释放），但养成习惯：
    ' 循环里反复 Set 大对象时，不释放会有内存/句柄压力
    ' ★ Set x = Nothing 只是"松开引用"，不等于关闭文件/程序；
    '   要关工作簿得 wb.Close，要退出 Excel 得 Application.Quit

    ' ---- 集合对象：Collection / Dictionary（属于"对象"，不是数组）----
    ' Dim c As Collection : Set c = New Collection
    ' c.Add 3.3, "VCC"                ' 值 + 可选键名
    ' MsgBox c("VCC")                 ' 按键取，重复键会报错
    '
    ' Dictionary 需要"工具 → 引用"勾选 Microsoft Scripting Runtime（或后期绑定 CreateObject）
    ' 好处：能用 Exists 判断键、能改值、能一次拿到 Keys/Items
    ' ★ 什么时候用：结果"按名字/编号存"，且数量不定 —— 比二维数组 + 循环查找清楚得多
End Sub
