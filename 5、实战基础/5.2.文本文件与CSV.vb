Option Explicit

' ============================================================
' 5.2 文本文件与 CSV：量产数据的读写主力
'     为什么不用 Excel 打开数据文件：
'       开工作簿慢、会弹窗、文件被占用、Excel 崩了脚本也崩；
'       几万行用文本流处理快一个数量级，也不依赖 Excel 状态
' ------------------------------------------------------------
' 和 Node 对照：
'   fs.readFileSync(p, "utf8").split(/\r?\n/)   →  Open For Input + Line Input 循环
'   fs.writeFileSync(p, s)                      →  Open For Output + Print #
'   fs.appendFileSync(p, s)                     →  Open For Append + Print #
'   fs.readdirSync(dir)                         →  Dir(path) 循环
' ============================================================
' ============================================================
' FreeFile：取一个"当前没人用"的文件号
'   Open / Line Input / Print # / Close 里的 #1、#2 就是文件号，
'   它在一个进程内共享，同一个号不能同时开两次 —— 所以别写死
' ============================================================
Sub DemoFreeFile()
    Dim f1 As Integer, f2 As Integer
    f1 = FreeFile
    f2 = FreeFile
    MsgBox "连续取两次：" & f1 & " / " & f2
    ' ★ 两个都是 1：FreeFile 只是"查一下哪个号空着"，并不真正占位，
    '   拿到号必须立刻 Open，中间别再调 FreeFile

    Dim path As String
    path = DemoPath()
    f1 = FreeFile
    Open path For Output As #f1
    MsgBox "占用 #" & f1 & " 之后再取：" & FreeFile      ' 通常变成 2

    Close #f1
    ' ★ Close 之后号被释放，下次又能取到它
    ' ★ 忘 Close 的后果：号一直被占着，循环里反复"开而不关"会把 1~255 耗完，
    '   报"文件过多"(67)；文件内容也可能还在缓冲区没落盘
    ' ★ # 不能省：Open ... As #fnum 里的 fnum 只是变量，井号必须手写
End Sub


Sub DemoFileRead()
    Dim path As String
    path = DemoPath()

    If Len(Dir$(path)) = 0 Then
        MsgBox "示例文件不存在：" & path & vbCrLf & "先运行 DemoFileWrite 造一个"
        Exit Sub
    End If

    Dim fnum As Integer, line As String, n As Long
    fnum = FreeFile                       ' ★ 别硬写 #1：嵌套调用会撞号
    Open path For Input As #fnum
    Do While Not EOF(fnum)                ' ★ 不判 EOF 会报"输入超出文件尾"(62)
        Line Input #fnum, line            ' 一次读一整行，已去掉行尾 CRLF
        n = n + 1
        If n <= 3 Then Debug.Print n & ": " & line   ' 看"立即窗口" Ctrl+G
    Loop
    Close #fnum
    MsgBox "共 " & n & " 行：" & path

    ' ---- 另外两种读法（知道有这回事就行）----
    ' Input #fnum, a, b, c       ' 专读"逗号分隔"的字段，会自己剥掉引号，
    '                            ' 但按变量类型解析，遇到脏数据直接报错，
    '                            ' 可控性不如自己 Split / SplitCsvLine
    ' Get #fnum, 1, buf          ' 二进制读，配 Open ... For Binary 用：
    '                            ' 适合定长记录、图片、或要精确控制字节（如去掉 BOM）
End Sub


Sub DemoFileWrite()
    Dim path As String, fnum As Integer
    path = DemoPath()

    fnum = FreeFile
    Open path For Output As #fnum         ' 覆盖写；要追加就换成 For Append
    ' ★ 自己拼好整行再 Print：Print #1, a, b 里的逗号是"制表位"不是分隔符，
    '   而且数值会被加上前导空格（正数留符号位），拼进 CSV 全是脏数据
    Print #fnum, "时间,引脚,电压,结果"
    Print #fnum, "2026-10-26 09:30:00,VCC," & CStr(3.31) & ",Pass"
    Print #fnum, "2026-10-26 09:30:01,GND," & CStr(0#) & ",Pass"
    Close #fnum                          ' ★ 不 Close：内容可能还在缓冲区，文件被占着

    MsgBox "已写入：" & path
End Sub


' ============================================================
' ★ 编码坑（中文必踩）：VBA 原生 Open 读写的是 ANSI（系统代码页）
'   UTF-8 文件用 For Input 打开 → 中文全是乱码，
'   而且不报错，直接拿乱码去匹配字符串，查半天
'   解决：读写 UTF-8 用 ADODB.Stream（系统自带，无需装东西）
' ============================================================
Function ReadUtf8(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                           ' 2 = adTypeText
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText(-1)            ' -1 = 全部读出
    st.Close
    Set st = Nothing
End Function


Sub WriteUtf8(ByVal path As String, ByVal text As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText text
    st.SaveToFile path, 2                 ' 2 = 覆盖；1 = 追加
    st.Close
    Set st = Nothing
    ' ★ 这样写出的 UTF-8 带 BOM，Excel 双击打开中文正常；
    '   若目标系统不认 BOM，就得手写字节流（Open For Binary + Put）去掉它
End Sub


' ============================================================
' 带引号的 CSV 不能用 Split：字段里出现逗号就会错列
'   规则：双引号包住的字段内的逗号属于内容；字段内的引号写成两个 ""
' ============================================================
Function SplitCsvLine(ByVal line As String) As Variant
    Dim res() As String, buf As String
    Dim inQuote As Boolean, ch As String
    Dim i As Long, n As Long

    ReDim res(0 To 0)
    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)

        If inQuote Then
            If ch = """" Then
                If Mid$(line, i + 1, 1) = """" Then
                    buf = buf & """"         ' "" → 一个字面引号
                    i = i + 1
                Else
                    inQuote = False          ' 引号区结束
                End If
            Else
                buf = buf & ch
            End If
        Else
            Select Case ch
                Case """": inQuote = True
                Case ",":                    ' 正常分隔
                    res(n) = buf
                    buf = ""
                    n = n + 1
                    ReDim Preserve res(0 To n)
                Case Else: buf = buf & ch
            End Select
        End If
    Next i

    res(n) = buf
    SplitCsvLine = res
End Function


Sub DemoParseCsv()
    Dim f As Variant
    ' 原始行：VCC,"3,310","已判""Pass"   （第 3 个字段里有一个转义引号）
    f = SplitCsvLine("VCC,""3,310"",""已判""""Pass""")   ' → VCC / 3,310 / 已判"Pass"
    MsgBox "段数：" & (UBound(f) + 1) & vbCrLf & _
           "第 2 段：" & f(1) & vbCrLf & _
           "第 3 段：" & f(2)
    ' ★ 段数不能假设：脏数据里少一列很常见，
    '   取值前先判 UBound(f) >= 期望下标，否则拿到的是空串而不是报错
End Sub


' ============================================================
' 遍历目录下的文件：Dir 有"状态"，用完要跑完或重开
' ============================================================
Sub DemoListFiles()
    Dim folder As String, name As String, n As Long
    folder = Environ$("TEMP")

    name = Dir$(folder & "\*.csv")
    Do While Len(name) > 0
        n = n + 1
        If n <= 5 Then Debug.Print n & ": " & name
        name = Dir$                        ' 不带参数 = 取下一个
    Loop
    MsgBox folder & " 下共 " & n & " 个 csv"
    ' ★ Dir 是全局状态：循环中途又调一次 Dir("其它路径") 会把游标打乱；
    '   递归遍历子目录要用 FSO 或自己写栈，别用 Dir 硬来
End Sub


' ============================================================
' 文件管理四件套：删 / 拷 / 改名 / 建目录
'   "每批结果自动归档、旧日志只留 30 天"这类需求全靠它们，
'   和正文读写是两回事，单独记
' ============================================================
Sub DemoFileManage()
    Dim dir As String, f1 As String, f2 As String
    dir = Environ$("TEMP")
    f1 = dir & "\ate_demo.csv"
    f2 = dir & "\ate_demo_bak.csv"

    If Len(Dir$(f1)) = 0 Then
        MsgBox "示例文件不存在，先跑 DemoFileWrite 造一个"
        Exit Sub
    End If

    ' ---- 复制：目标存在会直接覆盖，不提示 ----
    FileCopy f1, f2

    ' ---- 改名 / 移动：同一个语句，给新路径就等于"移动" ----
    ' Name f2 As dir & "\ate_demo_old.csv"
    ' ★ 目标已存在时 Name 报错 58（文件已存在），要先 Kill 或先判断

    ' ---- 删除：直接删，不进回收站，删了找不回 ----
    ' Kill f2
    ' ★ Kill 支持通配符：Kill dir & "\*.tmp"
    '   写通配符前先用 Dir 列一遍确认，别在正式数据目录上直接试
    ' ★ 文件不存在会报错 53，删除前先 Len(Dir$(f)) > 0 判一下

    ' ---- 文件属性查询 ----
    MsgBox "大小：" & FileLen(f1) & " 字节" & vbCrLf & _
           "最后修改：" & FileDateTime(f1)
    ' FileDateTime 返回 Date，"留 30 天"就靠它：
    ' If DateDiff("d", FileDateTime(f), Now) > 30 Then Kill f

    ' ---- 目录 ----
    ' MkDir dir & "\ATE_Log"        ' 已存在会报错 75（路径/文件访问错误）
    ' RmDir dir & "\ATE_Log"        ' 只能删空目录，里面有文件同样报 75
    ' ★ VBA 没有"递归删目录"，要自己 Dir 遍历 + Kill 清空，或者用 FSO.DeleteFolder
End Sub


' ============================================================
' Environ：读环境变量（临时目录、用户名、机器名、自定义路径）
'   JS 对照：process.env.TEMP / process.env.USERNAME
'   本文件里的 DemoPath / DemoListFiles 都用它拿到"临时目录"
' ============================================================
Sub DemoEnviron()
    MsgBox "临时目录：" & Environ$("TEMP")
    MsgBox "用户目录：" & Environ$("USERPROFILE")
    MsgBox "用户名：" & Environ$("USERNAME")
    MsgBox "机器名：" & Environ$("COMPUTERNAME")
    ' ★ 大小写不敏感：Environ$("temp") 一样能取到

    ' ---- 变量不存在时返回空串，不报错 ----
    Dim dir As String
    dir = Environ$("ATE_DATA_DIR")             ' 假设配了一个自定义变量
    If Len(dir) = 0 Then dir = Environ$("TEMP")
    MsgBox "数据目录：" & dir
    ' ★ 一定要判 Len：空串拼进路径会变成 "\ate.csv"，
    '   于是文件悄悄写到了当前盘符的根目录，找半天

    ' ---- 两个机台上的坑 ----
    ' 1) 服务账户 / 计划任务跑脚本时，TEMP 常指向 C:\Windows\Temp，
    '    跟你手动运行时不是同一个目录 —— "日志莫名不见了"多半是这个
    ' 2) 32 位和 64 位进程看到的系统变量可能不一致（SysWOW64 重定向）
    ' ★ 别用 Environ(1) 这种按序号取：顺序随系统变，没有任何保证
    ' ★ 也别拿环境变量存密码，同机任何进程都读得到
End Sub


' 示例文件路径：放临时目录，不动用户数据
' Environ$("TEMP") 的含义见上面的 DemoEnviron
Private Function DemoPath() As String
    Dim dir As String
    dir = Environ$("TEMP")
    If Len(dir) = 0 Then dir = "C:\Temp"       ' 兜底，避免拼出根目录
    ' ★ 这里的 DemoPath 是一个 String 变量，不是"调用自己"：
    '   VBA 没有 return，"给函数名赋值"就等于 return（见 3.1）
    '   不赋值就走到 End Function → 返回 ""（默认值），也不报错
    DemoPath = dir & "\ate_demo.csv"
End Function
