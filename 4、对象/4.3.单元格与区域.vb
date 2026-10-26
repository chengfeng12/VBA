Option Explicit

' ============================================================
' 4.3 单元格与区域（Range / Cells）
'     会改数据的语句都写在注释里，取消注释才会真的写入
' ============================================================
Sub DemoRange()
    Dim ws As Worksheet
    Dim rng As Range
    Set ws = ThisWorkbook.Worksheets(1)

    ' ---- 引用单元格的几种方式 ----
    Set rng = ws.Range("A1")                              ' 单个单元格
    Set rng = ws.Range("A1:C3")                           ' 一个矩形区域
    Set rng = ws.Cells(2, 3)                              ' 第 2 行第 3 列 → C2
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(3, 3))    ' A1:C3 的另一种写法
    Set rng = ws.Range("A1").Offset(1, 2)                 ' A1 向下 1 行、向右 2 列 → C2
    Set rng = ws.Range("A1").Resize(3, 2)                 ' 以 A1 为左上角，扩成 3 行 2 列

    MsgBox "Cells(2,3)      → " & ws.Cells(2, 3).Address & vbCrLf & _
           "Offset(1,2)     → " & ws.Range("A1").Offset(1, 2).Address & vbCrLf & _
           "Resize(3,2)     → " & ws.Range("A1").Resize(3, 2).Address
    ' ★ 记牢参数顺序：Cells(行, 列) 是 (row, column)
    '   而 Range("C2") 里的字符串是"先列后行"，两套读法容易打架

    ' ---- 读写值：Value / Value2 / Text / Formula 的区别 ----
    ' ws.Range("A1").Value = 3.3        ' 写入（最常用）
    ' MsgBox ws.Range("A1").Value       ' 读值：日期带时间、大数字可能丢精度
    ' MsgBox ws.Range("A1").Value2      ' 读值：更快，日期/货币保持底层数字，不还原类型
    ' MsgBox ws.Range("A1").Text        ' 屏幕上看到的字符串（按单元格格式）
    ' MsgBox ws.Range("A1").Formula     ' 公式文本，如 "=B1+C1"
    ' MsgBox ws.Range("A1").NumberFormat ' 单元格格式，如 "0.00"
    ' 结论：只取数就用 Value2，要"看到什么拿什么"就用 Text

    ' ---- 整行 / 整列 ----
    Set rng = ws.Rows(2)                    ' 第 2 行整行
    Set rng = ws.Columns("C")               ' C 列整列
    Set rng = ws.Range("A5").EntireRow      ' A5 所在整行
    Set rng = ws.Range("A5").EntireColumn   ' A5 所在整列

    ' ---- 定位数据边界 ----
    ' End(xlUp) 等价 Ctrl+↑，最常用的"找最后一行"
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    MsgBox "A 列最后一行：" & lastRow
    ' 四个方向：xlUp / xlDown / xlToLeft / xlToRight
    ' CurrentRegion ≈ Ctrl+A，取 A1 周围连续的数据块
    ' Set rng = ws.Range("A1").CurrentRegion
    ' UsedRange 是整张表"用过"的范围（可能残留已清空的空行空列，别完全信它）

    ' ---- 批量读写：一定走数组，逐格读写慢几十倍 ----
    ' 读：一块区域一次性读进二维数组，下标从 1 开始
    ' Dim arr As Variant
    ' arr = ws.Range("A1:C10").Value2
    ' MsgBox "第 2 行第 3 列 = " & arr(2, 3)      ' 注意是 (行, 列)

    ' 写：先建同样大小的二维数组，再一次性刷回去
    Dim out(1 To 2, 1 To 3) As Variant
    out(1, 1) = "VCC": out(1, 2) = 3.3: out(1, 3) = "Pass"
    out(2, 1) = "GND": out(2, 2) = 0#:  out(2, 3) = "Pass"
    ' ws.Range("E1").Resize(2, 3).Value = out    ' ← 取消注释才会真的写入
    MsgBox "准备写入 2 行 3 列，左上角 E1，实际范围 " & ws.Range("E1").Resize(2, 3).Address
    ' ★ 尺寸必须严丝合缝：2×3 的数组配 Resize(2, 3)，错一个就报"下标越界"或只写一部分
    ' ★ 一维数组不能直接往纵向区域灌：VBA 按"行"理解，容易只填一格/横向铺开；
    '   要竖着填就用二维数组 (1 To n, 1 To 1)，或套 WorksheetFunction.Transpose

    ' ---- 查找 ----
    ' Dim found As Range
    ' Set found = ws.Range("A1:A100").Find(What:="VCC", LookAt:=xlWhole)
    ' If Not found Is Nothing Then
    '     MsgBox "找到在 " & found.Address & "，第 " & found.Row & " 行"
    ' Else
    '     MsgBox "没找到"
    ' End If
    ' 注意：Find 会记住上次条件，复杂场景用 FindNext 循环，末尾记得重置
    ' ★ 漏第一个的经典坑：Find 从 After 所指的"下一个"格子开始找，默认 After 是区域首格，
    '   目标正好落在第一格时会绕一圈（表现成"没找到"或找到别处）→
    '   显式传 After:=区域最后一格，才是真正的完整扫描
    ' ★ 参数会"粘住"：这次传了 LookAt:=xlWhole，下次不写也仍是整格匹配；
    '   FindNext 循环不会自动换参数，找不同条件前先重设一次

    ' ---- 复制 / 粘贴 / 清除 ----
    ' ws.Range("A1").Copy Destination:=ws.Range("D1")
    ' ws.Range("A1:C3").Copy
    ' ws.Range("D1").PasteSpecial Paste:=xlPasteValues          ' 只粘值
    ' ws.Range("D1").PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd  ' 粘值并累加
    ' Application.CutCopyMode = False                            ' 清掉"蚂蚁线"
    ' ws.Range("A1:C3").ClearContents     ' 只清内容，保留格式
    ' ws.Range("A1:C3").Clear             ' 内容 + 格式 + 批注全清
    ' ws.Range("A1:C3").ClearFormats      ' 只清格式

    ' ---- 其他常用 ----
    ' Intersect(ws.Range("A1:C3"), ws.Range("B2:D5"))   ' 交集 → B2:C3
    ' Union(ws.Range("A1"), ws.Range("C3"))             ' 并集
    ' ws.Range("A1").Comment.Text "备注"                 ' 加批注
    ' ★ Intersect 没有交集时返回 Nothing，套 Not ... Is Nothing 判断再用

    ' ---- 合并单元格：写入前先取 MergeArea ----
    ' If ws.Range("A1").MergeCells Then
    '     MsgBox "A1 属于合并区：" & ws.Range("A1").MergeArea.Address
    ' End If
    ' ★ 直接给合并区的非左上角写值，值会"消失"或不显示，先 .MergeArea 再写
    ' ★ 合并区的 .Value 只有左上角有值，其余格读出来是 Empty

    ' ---- 错误值：读之前先判 ----
    ' If IsError(ws.Range("A1").Value) Then MsgBox "这格是 #N/A 之类的错误值"
    ' ★ 不判就参与计算/比较，会直接把整个表达式变成一个错误值（和 JS 的 NaN 一样会传染）
End Sub


' ============================================================
' 找某列最后一个非空行（空表时返回 0，比直接 End(xlUp) 安全）
' ============================================================
Function LastRowOf(ByVal ws As Worksheet, ByVal col As Long) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If r = 1 And Len(ws.Cells(1, col).Value & "") = 0 Then
        LastRowOf = 0            ' 整列为空，End(xlUp) 会停在 1，容易误判成"有 1 行"
    Else
        LastRowOf = r
    End If
End Function
