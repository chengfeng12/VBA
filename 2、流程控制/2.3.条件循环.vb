Option Explicit

' ============================================================
' 2.3 条件循环：次数未知、靠条件停的时候用它
'     Do While 条件 ... Loop     先判断，可能一次都不执行
'     Do ... Loop While 条件     先执行，至少执行一次
'     Do Until 条件 ... Loop     Until = 条件"成立"时停（和 While 方向相反）
'   四种组合都合法：判断在前/在后 × While/Until
' ============================================================
Sub DemoDoLoop()

    ' ---- 前测：条件一上来就不成立 → 一次都不执行 ----
    Dim n As Long
    n = 0
    Do While n < 3
        n = n + 1
    Loop
    MsgBox "前测 Do While → n = " & n                 ' 3

    ' ---- 后测：至少执行一次 ----
    Dim m As Long
    m = 100
    Do
        m = m + 1
    Loop While m < 3
    MsgBox "后测 Do Loop While → m = " & m           ' 101
    ' ★ 差的那一次就在这里：本来 100 已经不满足条件，但因为"先执行后判断"，
    '   还是会加一次。写"至少要试一次"的逻辑（如首次连接设备）才用后测

    ' ---- Until：条件为真时退出，判定方向和 While 相反 ----
    Dim k As Long
    k = 0
    Do Until k >= 3
        k = k + 1
    Loop
    MsgBox "Do Until → k = " & k                      ' 3
    ' 记忆：While 是"成立就继续"，Until 是"成立就停"

    ' ---- 死循环防护：外部条件不可控时必须加最大次数兜底 ----
    Dim retry As Long, ok As Boolean
    retry = 0
    Do While (Not ok) And retry < 5                   ' 最多重试 5 次
        retry = retry + 1
        ok = (retry >= 3)                             ' 假装第 3 次才成功
    Loop
    MsgBox "重试 " & retry & " 次，结果 " & IIf(ok, "Pass", "Fail")
    ' ★ 和硬件、外部文件打交道的循环，条件里一定要有"次数上限"，
    '   否则设备不返回应答就是整片界面卡死

    ' ---- Exit Do：从中间跳出，比把一个长条件塞进 While 好读 ----
    Dim t As Long
    t = 0
    Do
        t = t + 1
        If t = 4 Then Exit Do
    Loop
    MsgBox "Exit Do → t = " & t

    ' ---- While ... Wend：老写法，只有前测、没有 Exit，了解即可 ----
    ' Dim w As Long
    ' While w < 3
    '     w = w + 1
    ' Wend
    ' 结论：新代码一律用 Do While，别用 While...Wend

    ' ---- 调试提示 ----
    ' 循环写到一半卡死，按 Ctrl + Break（或 Esc）中断，
    ' 再把光标放到变量上、或用"立即窗口"（Ctrl + G）打印值看停在哪一轮
End Sub
