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
    ' ★ ThisWorkbook 是"代码所在的工作簿"；ActiveWorkbook 是"用户当前激活的那个"，
    '   用户切一下窗口两者就不是一个文件了 —— 脚本里一律用 ThisWorkbook

    ' ---- CodeName：用户改标签页名字也不会失效 ----
    MsgBox "代号：" & ws.CodeName & "，显示名：" & ws.Name
    ' 代码里写 Worksheets("Sheet1") 依赖显示名，用户一改名就崩；
    ' 用对象代号（工程里直接写 Sheet1.Range(...)）或用 CodeName 反查，才稳

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
    ' ★ 一次删多张别用 For Each：集合边遍历边删会跳表或报错，
    '   倒序按下标删才安全：
    '   Dim i As Long
    '   For i = wb.Worksheets.Count To 1 Step -1
    '       If 条件 Then wb.Worksheets(i).Delete
    '   Next
    ' ★ 唯一一张工作表删不掉，至少留一张可见表

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
'   ★ 代价：任何原因取不到（名字含非法字符、工作簿结构被保护）都会算"不存在"
'      —— 拿它做"要不要新建"的前置判断够用，别当权限校验
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


' ============================================================
' 提速三件套：ScreenUpdating / Calculation / EnableEvents
'   目录里只写了"用完必须恢复"，这里给出完整写法：
'   中途出错也要恢复 → 用 2.4 的 Cleanup 模式，恢复正常路径也会掉进去
'   收益：批量写几万行时，不关这三项会慢十倍以上
' ============================================================
Sub DemoFastMode()
    Dim prevCalc As XlCalculation
    Dim prevEvents As Boolean

    prevCalc = Application.Calculation            ' ★ 先存原状态，别假设一定是"自动"
    prevEvents = Application.EnableEvents

    On Error GoTo Cleanup
    Application.ScreenUpdating = False            ' 屏幕不刷新
    Application.Calculation = xlCalculationManual ' 公式不自动重算
    Application.EnableEvents = False              ' 不触发 Worksheet_Change 等事件

    ' ...这里放批量读写...
    MsgBox "静默模式：屏幕不刷新、公式不重算、事件不触发"

Cleanup:
    On Error Resume Next                          ' 恢复过程本身不能再引发跳转
    Application.ScreenUpdating = True
    Application.Calculation = prevCalc            ' ★ 恢复原值，别硬写 xlCalculationAutomatic
    Application.EnableEvents = prevEvents
    ' ★ 忘了恢复的典型症状：界面一片灰不刷新、公式改了不重算、
    '   连自定义按钮都失灵 —— 只能重启 Excel
End Sub
