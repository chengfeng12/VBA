Option Explicit

' ============================================================
' 3.3 错误处理：On Error GoTo / Err 对象 / Resume / Err.Raise
'     VBA 没有 try...catch，只有"出错就跳到某个标签"
' ------------------------------------------------------------
'     On Error GoTo 标签     出错跳到标签（主力用法）
'     On Error Resume Next   出错就跳过这一句，接着往下跑
'     On Error GoTo 0        关闭错误处理，恢复"出错就弹框"
' ============================================================
Sub DemoErrorHandling()
    On Error GoTo Failed

    MsgBox "安全转换：[" & "22.5" & "] → " & SafeDouble("22.5")
    MsgBox "安全转换：[" & " 1,234.5 " & "] → " & SafeDouble(" 1,234.5 ")

    MsgBox "下面这行会抛错，然后跳到 Failed" & vbCrLf & "（" & "abc" & " 不是数字）"
    MsgBox "安全转换：[" & "abc" & "] → " & SafeDouble("abc")    ' ← 在这里出错

    Exit Sub                            ' ★ 必须写：否则会顺着往下掉进错误处理块

Failed:
    MsgBox "出错编号：" & Err.Number & vbCrLf & _
           "描述：" & Err.Description & vbCrLf & _
           "来源：" & Err.Source
    ' Resume 的三种收尾方式
    '   Resume          回到出错的那一行重试（问题修好了才敢用，否则死循环）
    '   Resume Next     跳过出错语句，从下一行继续（批处理里"容忍个别失败"）
    '   Resume 标签     跳到指定位置
    Err.Clear                           ' 处理完手动清掉，避免残留到下一次判断
End Sub


' ============================================================
' On Error Resume Next：出错就跳过，适合"允许个别失败"的批处理
' ★ 最大的坑：用完忘了 On Error GoTo 0，后面几百行全在静默失败
' ============================================================
Sub DemoResumeNext()
    Dim arr As Variant, i As Long, okCnt As Long, total As Long
    arr = Array("3.3", "abc", "0.01")
    total = UBound(arr) - LBound(arr) + 1

    On Error Resume Next                ' 从这里开始"忽略错误"
    For i = LBound(arr) To UBound(arr)
        Err.Clear                       ' ★ 每轮先清：否则读到的是上一轮的旧错误
        Debug.Print SafeDouble(arr(i))  ' 结果看"立即窗口"（Ctrl + G）
        If Err.Number = 0 Then okCnt = okCnt + 1
    Next
    On Error GoTo 0                     ' ★ 用完立刻关掉

    MsgBox "成功 " & okCnt & " 个，失败 " & (total - okCnt) & " 个"
End Sub


' ============================================================
' 往上抛错误：Err.Raise 编号, 来源, 描述
' 编号规则：vbObjectError + 自己定的一段（1~65535 会和系统错误号撞）
' ============================================================
Function SafeDouble(ByVal v As Variant) As Double
    Dim clean As String

    If VarType(v) = vbString Then
        clean = Trim$(v)
        clean = Replace(clean, ",", "")         ' 去千分位
        If Not IsNumeric(clean) Then GoTo Bad
        SafeDouble = Val(clean)                 ' 用 Val 避开 CDbl 的区域设置坑（1.3 / 1.6）
    Else
        On Error GoTo Bad
        SafeDouble = CDbl(v)
    End If
    Exit Function

Bad:
    Err.Raise vbObjectError + 1001, "SafeDouble", "不是合法数字：[" & CStr(v) & "]"
End Function
