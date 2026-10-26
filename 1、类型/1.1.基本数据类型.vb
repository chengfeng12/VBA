Option Explicit

' ============================================================
' 1.1 基本数据类型：Dim 变量名 As 类型
'     声明成什么类型，就只能装什么（VBA 会自动做隐式转换）
' ============================================================
Sub DemoBasicTypes()
    ' ---- 整数类 ----
    ' 简写对照：Integer 用 %、Long 用 &、Byte 没有符号
    '   Dim i%     就是 Dim i As Integer
    '   Dim l&     就是 Dim l As Long
    ' ★ 符号不能和 As 混用：Dim i% As Integer 是语法错误（详见 1.2）
    ' ★ 日常一律写 As 类型，简写留给字面量后缀（30000& * 2 防溢出，见 1.2）
    Dim b As Byte       ' 0 ~ 255                        1 字节
    ' 例：b = 200            引脚编号、颜色分量（0~255）都够；再大就"溢出"
    Dim i As Integer    ' -32,768 ~ 32,767                2 字节
    ' 例：i = 30000          可以；改 40000 直接报"溢出"；30000 * 2 也溢出
    Dim l As Long       ' -21亿 ~ 21亿                    4 字节（循环计数器首选）
    ' 例：l = 479001600      12! 阶乘；For i = 1 To 100000 一律用 Long
    ' Dim ll As LongLong  ' 8 字节，仅 64 位 Office 支持；32 位会报"用户定义类型未定义"
    ' 例：ll = 3000000000    35 亿，Long 装不下，只能靠它

    ' ---- 小数类 ----
    ' 简写对照：Single 用 !、Double 用 #、Currency 用 @
    '   Dim s!     就是 Dim s As Single
    '   Dim d#     就是 Dim d As Double        ← 和 #日期# 靠上下文区分，不冲突
    '   Dim c@     就是 Dim c As Currency
    Dim s As Single     ' 单精度，约 7 位有效数字          4 字节
    ' 例：s = 3.1415927      第 8 位起就不准了，判限别用 Single
    Dim d As Double     ' 双精度，约 15 位有效数字         8 字节（工程计算首选）
    ' 例：d = 3.3            实测电压都用它；注意 0.1 + 0.2 <> 0.3（见 1.7）
    Dim c As Currency   ' 定点数，固定 4 位小数            8 字节（金额用，无浮点误差）
    ' 例：c = 12.3456        第 5 位会被舍入；0.1 + 0.2 精确等于 0.3

    ' ---- 文本 / 逻辑 / 日期 / 对象 ----
    ' 简写对照：String 用 $（Dim str$），Boolean / Date / Object 没有符号
    ' ★ 但 $ 更常见的用处是"函数简写"：Trim$/Left$/Mid$ 返回 String 而非 Variant，
    '   文本处理优先用带 $ 的版本（见 1.2 与 5.1）
    Dim str As String   ' 字符串，变长
    ' 例：str = "VCC_3V3"    引脚名、日志行；拼接用 & 不用 +（见下）
    Dim flag As Boolean ' True / False（注意：VBA 里 True = -1，False = 0）
    ' 例：flag = (3.3 >= 3)  结果是 True；CInt(flag) 才是 -1，
    '                        直接拼字符串打印仍然是 "True"
    Dim dt As Date      ' 日期时间，范围 100/1/1 ~ 9999/12/31
    ' 例：dt = #10/26/2026 9:30:00 AM#   字面量用 # 包；
    '                        顺序按系统区域解析，跨机器稳妥用 DateSerial
    Dim obj As Object   ' 任意对象引用：Workbook / Worksheet / Range ...
    ' 例：Set obj = ThisWorkbook        对象必须用 Set（见 4.1）

    ' ---- 赋值 ----
    b = 200
    i = 30000
    l = 100000            ' 如果 i 也赋 100000，会报"溢出"，因为 Integer 装不下
    s = 3.14
    d = 3.14159265358979
    c = 1234.5678
    str = "ATE 自动化"
    ' ★ 字符串拼接一律用 &，别用 +：
    '     "1" & 1   → "11"        "1" + 1    → 2（退化成加法）
    '     Null & "a" → "a"        Null + "a" → Null（污染整个表达式）
    '   & 两边必须留空格：s&i 会被当成"s 是 Long 型"（& 是 Long 的类型声明字符）
    '   VBA 没有模板字符串，多段拼接只能手写 &，换行用 vbCrLf / vbTab
    flag = True
    dt = DateSerial(2026, 10, 22) + TimeSerial(9, 30, 0) ' 也可以用 #10/22/2026 9:30# 这种字面量
    Set obj = ThisWorkbook ' 对象必须用 Set 赋值

    ' ★ Format 里 m 是"月"、n 才是"分"：
    '   写 "hh:mm:ss" 会打出小时 + 月份 + 秒，看着像时间但是错的；
    '   24 小时制要写 "yyyy-mm-dd hh:nn:ss"
    MsgBox "Byte: " & b & vbCrLf & _
           "Integer: " & i & vbCrLf & _
           "Long: " & l & vbCrLf & _
           "Single: " & s & vbCrLf & _
           "Double: " & d & vbCrLf & _
           "Currency: " & c & vbCrLf & _
           "String: " & str & vbCrLf & _
           "Boolean: " & flag & vbCrLf & _
           "Date: " & Format(dt, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
           "Object: " & obj.Name
End Sub
