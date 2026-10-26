Option Explicit

' ============================================================
' 3.3 错误处理：On Error GoTo / Err 对象 / Resume / Err.Raise
'     VBA 没有 try...catch，只有"出错就跳到某个标签"
'
' ★ 和 JS 最大的差别在"作用域"，不是语法：
'   JS 的 try 是块级的，可以只包三五行，出错就 catch、正常就走 finally；
'   VBA 的 On Error 是【过程级开关】，一设就管到这个过程结束（或下次重设），
'   包不住"某一段"代码 —— 想局部容错，只能把那段拆成独立过程再调用。
' ★ 也没有 finally：收尾位置就是错误处理器的标签（见 2.4 的 CleanupDemo）
' ------------------------------------------------------------
'     On Error GoTo 标签     出错跳到标签（主力用法）
'     On Error Resume Next   出错就跳过这一句，接着往下跑，不会造成阻塞
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
    ' ★ 处理器里再出错 → 又跳回 Failed 本身，可能无限刷屏；
    '   所以真正的收尾块（如 2.4 的 CleanupDemo）第一行都是 On Error Resume Next
    ' ★ 在没有任何错误的情况下执行 Resume 会报"Resume without error"（错误号 20），
    '   所以 Resume 只能待在错误处理块里，不能当普通跳转用
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
' Err 的四个实战坑（踩过才记得住）
'   1) Err 是全局单例：子过程里 Err.Clear 一下，外层就再也读不到错误了
'      → 判断错误要"出错后立刻判断"，中间别插别的过程调用
'   2) 错误会往上冒泡：子过程自己没有 On Error，错误交给调用方处理；
'      若调用方是 On Error Resume Next，就从"调用子过程那一行"的下一行继续，
'      子过程剩下的代码不再执行 —— 表现是"没跑完，但也没报错"
'   3) Err.Clear 只清 Err 对象的属性，错误仍处于"活动"状态；
'      On Error GoTo -1 才是彻底退出错误状态（VBA7 语法，日常混用没问题，知道区别即可）
'   4) On Error 是过程级的：过程结束或 On Error GoTo 0 之后自动失效，不会传给子过程
' ============================================================


' ============================================================
' 往上抛错误：Err.Raise 编号, 来源, 描述
'   编号规则：vbObjectError + 自己定的一段（1~65535 会和系统错误号撞）
'
'   vbObjectError 是什么：一个固定的负数常量（-2147221504），
'   专门留给"应用程序自定义错误"的号段起点。加 1001 就是你的第 1001 号业务错误，
'   保证不会和系统错误（如 13 类型不匹配、53 文件不存在）撞号。
'   写 vbObjectError + 1001 比直接写 -2147220503 好读一万倍
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


' ============================================================
' 抛出去的错误要有人接：Err.Raise 之后，调用方靠 Err.Number 分辨
'   判断一律用编号，别用 Description —— 系统错误的描述文字随 Office 语言版本变
'   自定义错误号 = vbObjectError + 1001，判断时要减回去（或像这里用常量对齐）
' ============================================================
Sub DemoCatchCustomError()
    Const ERR_NOT_NUMBER As Long = vbObjectError + 1001      ' 必须和 SafeDouble 里一致

    Dim v As Variant, d As Double
    On Error GoTo Failed

    v = "abc"
    d = SafeDouble(v)
    MsgBox "转换成功：" & d
    Exit Sub

Failed:
    Select Case Err.Number
        Case ERR_NOT_NUMBER
            MsgBox "接住自己抛的业务错误：" & Err.Description
        Case 13                                            ' 13 = 类型不匹配（系统错误号）
            MsgBox "类型不对：" & Err.Description
        Case Else
            MsgBox "未预期的错误 " & Err.Number & "（来自 " & Err.Source & "）"
    End Select
    Err.Clear
    ' ★ 这里刻意不写 Resume：错误已经用人话报告给用户了，
    '   再从出错处重试只会重复报错。想重试就写 Resume，想跳过就写 Resume Next
End Sub
