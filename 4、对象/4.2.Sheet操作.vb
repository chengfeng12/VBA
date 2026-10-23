Option Explicit

' ============================================================
' 4.2 Sheet 操作
'     Worksheets 只含普通工作表；Sheets 还包含图表工作表、宏表
'     增删改一类操作写在注释里，需要时自己取消注释
' ============================================================
Sub DemoSheets()
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ThisWorkbook

    ' ---- 引用工作表 ----
    Set ws = wb.Worksheets(1)                  ' 按序号（从 1 开始）
    ' Set ws = wb.Worksheets("Sheet1")         ' 按名称（改过名就用这个）
    ' Set ws = wb.Sheets(1)                    ' 含图表工作表时也能取

    MsgBox "共 " & wb.Worksheets.Count & " 张表，第 1 张叫：" & wb.Worksheets(1).Name

    ' ---- 遍历所有表 ----
    Dim names As String
    For Each ws In wb.Worksheets
        names = names & ws.Name & "  "
    Next
    MsgBox "所有表：" & names

    ' ---- 判断表是否存在（函数见本模块下方）----
    MsgBox "Sheet1 是否存在：" & SheetExists("Sheet1")

    ' ---- 新增 ----
    ' Dim newWs As Worksheet
    ' Set newWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ' newWs.Name = "ATE_Report"                 ' 名字重复会报错，先判断存在

    ' ---- 重命名 ----
    ' ws.Name = "数据源"

    ' ---- 复制 / 移动 ----
    ' ws.Copy After:=wb.Worksheets(wb.Worksheets.Count)   ' 复制到本工作簿末尾
    ' ws.Copy Before:=wb.Worksheets(1)                    ' 复制到最前面
    ' ws.Copy                                             ' 不带参数 → 新建一个工作簿装它
    ' ws.Move After:=wb.Worksheets(1)                     ' 移动顺序

    ' ---- 删除 ----
    ' Application.DisplayAlerts = False        ' 关掉"确认删除"弹窗
    ' wb.Worksheets("ATE_Report").Delete
    ' Application.DisplayAlerts = True         ' ★ 一定要恢复，否则后续操作都不提示

    ' ---- 隐藏 ----
    ' ws.Visible = xlSheetHidden               ' 普通隐藏，用户可右键取消
    ' ws.Visible = xlSheetVeryHidden           ' 深度隐藏，只能改 VBE 属性
    ' ws.Visible = xlSheetVisible              ' 恢复显示

    ' ---- 激活 / 选中 ----
    ' ws.Activate                              ' 激活（会切换当前表）
    ' ws.Select                                ' 选中
    ' 注意：能不激活就不激活，Activate 会拖慢速度、还会改变用户当前视线

    ' ---- 保护 / 取消保护 ----
    ' ws.Protect Password:="123", DrawingObjects:=True, Contents:=True
    ' ws.Unprotect Password:="123"
    MsgBox "是否受保护：" & ws.ProtectContents

    ' ---- 常用属性 ----
    MsgBox "使用区域：" & ws.UsedRange.Address & vbCrLf & _
           "索引位置：" & ws.Index & vbCrLf & _
           "可见性代码：" & ws.Visible & "（-1 可见，0 隐藏，2 深度隐藏）"
End Sub


' ============================================================
' 判断工作表是否存在
'   用 On Error 是 VBA 里最省事的写法：取不到就说明不存在
'   Optional 参数省掉时默认查 ThisWorkbook
' ============================================================
Function SheetExists(ByVal sheetName As String, Optional ByRef wb As Workbook) As Boolean
    Dim target As Workbook
    Dim tmp As Worksheet

    If wb Is Nothing Then
        Set target = ThisWorkbook
    Else
        Set target = wb
    End If

    On Error Resume Next
    Set tmp = target.Worksheets(sheetName)
    On Error GoTo 0

    SheetExists = Not (tmp Is Nothing)
    Set tmp = Nothing
    Set target = Nothing
End Function
