Option Explicit

' ============================================================
' 1.1 基本数据类型：Dim 变量名 As 类型
'     声明成什么类型，就只能装什么（VBA 会自动做隐式转换）
' ============================================================
Sub DemoBasicTypes()
    ' ---- 整数类 ----
    Dim b As Byte       ' 0 ~ 255                        1 字节
    Dim i As Integer    ' -32,768 ~ 32,767                2 字节
    Dim l As Long       ' -21亿 ~ 21亿                    4 字节（循环计数器首选）
    ' Dim ll As LongLong  ' 8 字节，仅 64 位 Office 支持；32 位会报"用户定义类型未定义"

    ' ---- 小数类 ----
    Dim s As Single     ' 单精度，约 7 位有效数字          4 字节
    Dim d As Double     ' 双精度，约 15 位有效数字         8 字节（工程计算首选）
    Dim c As Currency   ' 定点数，固定 4 位小数            8 字节（金额用，无浮点误差）

    ' ---- 文本 / 逻辑 / 日期 / 对象 ----
    Dim str As String   ' 字符串，变长
    Dim flag As Boolean ' True / False（注意：VBA 里 True = -1，False = 0）
    Dim dt As Date      ' 日期时间，范围 100/1/1 ~ 9999/12/31
    Dim obj As Object   ' 任意对象引用：Workbook / Worksheet / Range ...

    ' ---- 赋值 ----
    b = 200
    i = 30000
    l = 100000            ' 如果 i 也赋 100000，会报"溢出"，因为 Integer 装不下
    s = 3.14
    d = 3.14159265358979
    c = 1234.5678
    str = "ATE 自动化"
    flag = True
    dt = DateSerial(2026, 10, 22) + TimeSerial(9, 30, 0) ' 也可以用 #10/22/2026 9:30# 这种字面量
    Set obj = ThisWorkbook ' 对象必须用 Set 赋值

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
