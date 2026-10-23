Option Explicit

' ============================================================
' 2.4 跳转与退出：提前离开，减少嵌套
'     Exit Sub / Exit Function   离开整个过程
'     Exit For / Exit Do         离开当前这一层循环（2.2 / 2.3 已见）
'     GoTo 标签                  跳到同一过程内的另一个位置
'     GoTo 只能在本过程内跳，不能跳进循环内部、不能跳到别的模块
' ============================================================
Sub DemoJump()

    ' ---- 卫语句（guard clause）：不合格就提前 Exit Sub ----
    ' 两个样例对比看效果：第一个走完全程，第二个在卫语句处就结束
    CheckPin 3.3
    CheckPin 4.2

    ' ---- GoTo 的正当用法：统一收尾 ----
    CleanupDemo

    ' ---- 结论 ----
    ' 分支用 If / Select Case，循环提前结束用 Exit For / Exit Do，
    ' 过程提前结束用 Exit Sub / Exit Function；
    ' GoTo 只留给"跳到末尾做善后"这一种场景
End Sub


' ============================================================
' 卫语句示范：先挡掉不合格的情况，主逻辑不用层层缩进
' 不合法 → 直接 Exit Sub，后面的语句根本不执行
' ============================================================
Sub CheckPin(ByVal v As Double)
    If Not Judge(v, 3#, 3.6) Then
        MsgBox v & " V 超限 → Fail，本次测试到此结束"
        Exit Sub
    End If
    MsgBox v & " V 合格 → 继续下一项测试"
End Sub


Sub CleanupDemo()
    ' 中间任何一步出错都跳到 Cleanup 做善后，
    ' 清理代码只写一份，不会因为新增分支而漏掉（第 4 章的提速开关就靠它恢复）
    Dim state As String
    state = "未开始"
    On Error GoTo Cleanup                     ' 3.3 会细讲 On Error

    state = "处理中"
    ' ... 真正干活的代码 ...

Cleanup:
    On Error Resume Next
    ' 这里做善后：恢复 ScreenUpdating、关闭文件、写日志 ...
    MsgBox "状态：" & state & " → 已收尾"
    ' ★ 标签后面带冒号、顶格写；标签名在同一个过程内不能重名
    ' ★ 标签只是"位置标记"：程序正常往下走时也会穿过它，
    '   所以正常路径要么 Exit Sub 提前离开，要么像上面这样让收尾能安全重入
End Sub


' ============================================================
' 判定函数：返回值是 Boolean，默认值 False，
' 所以"提前返回"时什么都不用写，直接 Exit Function
' ============================================================
Function Judge(ByVal v As Double, ByVal lo As Double, ByVal hi As Double) As Boolean
    If v < lo Then Exit Function
    If v > hi Then Exit Function
    Judge = True
End Function
