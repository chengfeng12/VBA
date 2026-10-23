Option Explicit

' ============================================================
' 4.4 表格（ListObject，即 Excel 里的「超级表」）
'     核心价值：结构化引用 —— 用列名代替单元格地址，
'     插入行/列、挪位置都不会让公式和代码失效
'     会改结构的语句都在注释里
' ============================================================
Sub DemoTable()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Set ws = ThisWorkbook.Worksheets(1)

    ' ---- 建表：把 A1:C10 这块钱变成超级表 ----
    ' Set tbl = ws.ListObjects.Add(SourceType:=xlSrcRange, _
    '                              Source:=ws.Range("A1:C10"), _
    '                              XlListObjectHasHeaders:=xlYes)
    ' tbl.Name = "tblATE"                     ' 表名要唯一，且不能和单元格名冲突
    ' tbl.TableStyle = "TableStyleMedium9"    ' 换个样式

    ' ---- 取已有的表 ----
    If ws.ListObjects.Count = 0 Then
        MsgBox "这张表上还没有超级表。" & vbCrLf & _
               "想试的话，先选中一块带标题的数据，按 Ctrl+T 转成表格，再运行本过程。"
        Exit Sub
    End If

    Set tbl = ws.ListObjects(1)
    ' Set tbl = ws.ListObjects("tblATE")      ' 也可以按表名取

    MsgBox "表名：" & tbl.Name & vbCrLf & _
           "数据行数：" & tbl.ListRows.Count & vbCrLf & _
           "列数：" & tbl.ListColumns.Count & vbCrLf & _
           "整个表的范围：" & tbl.Range.Address

    ' ---- 结构化引用的精髓：按列名取值 ----
    ' 老写法 ws.Range("B2")：插一列就全错
    ' 新写法 tbl.ListColumns("电压").DataBodyRange(1)：永远指向"电压"列第 1 行数据
    ' MsgBox "电压列第 1 行 = " & tbl.ListColumns("电压").DataBodyRange(1)

    ' ---- 四个常用范围 ----
    ' tbl.Range            整个表（含表头和汇总行）
    ' tbl.HeaderRowRange   表头行
    ' tbl.DataBodyRange    数据区（不含表头、不含汇总行）← 最常用
    ' tbl.TotalsRowRange   汇总行（前提是 ShowTotals = True）

    ' ---- 增行 / 删行 ----
    ' Dim newRow As ListRow
    ' Set newRow = tbl.ListRows.Add                     ' 末尾追加一行
    ' newRow.Range(1, 1).Value = "VCC"                  ' 第 1 列
    ' newRow.Range(1, 2).Value = 3.3                    ' 第 2 列
    ' tbl.ListRows(tbl.ListRows.Count).Delete           ' 删最后一行
    ' 好处：表格的大小会自动扩展，不需要手动改区域

    ' ---- 按列名批量读写（性能最优）----
    ' Dim v As Variant
    ' v = tbl.ListColumns("电压").DataBodyRange.Value2        ' 整列读成二维数组
    ' tbl.ListColumns("结果").DataBodyRange.Value = "Pass"    ' 整列统一赋值
    ' tbl.DataBodyRange.Value2                                ' 全表读成二维数组

    ' ---- 筛选 ----
    ' tbl.Range.AutoFilter Field:=3, Criteria1:="Fail"    ' Field 是"第几列"（1 开始）
    ' tbl.Range.AutoFilter                                ' 不带参数 = 取消筛选
    ' 注意 Field 的序号是相对表格范围的，不是工作表的列号

    ' ---- 排序 ----
    ' With tbl.Sort
    '     .SortFields.Clear
    '     .SortFields.Add Key:=tbl.ListColumns("电压").DataBodyRange, _
    '                     SortOn:=xlSortOnValues, _
    '                     Order:=xlAscending, _
    '                     DataOption:=xlSortNormal
    '     .Header = xlYes
    '     .Apply
    ' End With

    ' ---- 汇总行 ----
    ' tbl.ShowTotals = True
    ' tbl.ListColumns("电压").TotalsCalculation = xlTotalsCalculationAverage
    ' 常用值：xlTotalsCalculationSum / Average / Count / Max / Min / StdDev

    ' ---- 转回普通区域 ----
    ' tbl.Unlist
    ' 注意：Unlist 之后结构化引用全部失效，代码里所有 tbl.xxx 都得改回 Range

    ' ---- 按表名找表（跨表操作时有用）----
    ' Dim found As ListObject
    ' Set found = FindTable(ThisWorkbook, "tblATE")
End Sub


' ============================================================
' 在整个工作簿里按表名找超级表（找不到返回 Nothing）
' ============================================================
Function FindTable(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim tbl As ListObject

    For Each ws In wb.Worksheets
        For Each tbl In ws.ListObjects
            If StrComp(tbl.Name, tableName, vbTextCompare) = 0 Then
                Set FindTable = tbl
                Exit Function
            End If
        Next
    Next
End Function
