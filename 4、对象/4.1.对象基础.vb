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

    ' ---- With：省掉重复的对象前缀 ----
    With ws.Range("B2")
        '.Value = "VCC"             ' 真写入，先注释
        MsgBox "地址：" & .Address & vbCrLf & _
               "行：" & .Row & "，列：" & .Column
    End With
    ' With 块里以 . 开头的是 With 后面那个对象；不写 . 就是普通变量/过程

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
End Sub
