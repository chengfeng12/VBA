Option Explicit

' ============================================================
' 1.2 简写形式：类型声明字符（type declaration character）
'     每种类型配一个符号，声明时、写字面量时都能用
'       %  Integer      &  Long        ^  LongLong（64 位）
'       !  Single       #  Double      @  Currency      $  String
'     注意：符号不能和 As 混用，Dim i% As Integer 是语法错误
' ============================================================
Sub DemoShortForm()
    ' ---- 声明简写：Dim 名字+符号，完全等价于 Dim 名字 As 类型 ----
    Dim i%              ' = Dim i As Integer
    Dim l&              ' = Dim l As Long
    Dim s!              ' = Dim s As Single
    Dim d#              ' = Dim d As Double
    Dim c@              ' = Dim c As Currency
    Dim str$            ' = Dim str As String
    ' Dim ll^           ' = Dim ll As LongLong，仅 64 位 Office 支持

    ' 数组也能这样写
    Dim arr%(1 To 3)    ' = Dim arr(1 To 3) As Integer

    ' ---- 字面量简写：给数字加后缀，强制这个常量的类型 ----
    i% = 100
    l& = 100000&        ' & 后缀 → 这个 100000 是 Long 字面量
    s! = 3.14!          ' ! 后缀 → Single
    d# = 3.14159265358979# ' # 后缀 → Double（和 #日期# 靠上下文区分，不冲突）
    c@ = 1234.5678@     ' @ 后缀 → Currency
    str$ = "ATE 自动化"  ' 字符串本身就是 String，不用加 $（加了也合法，纯冗余）
    arr%(1) = 1

    ' 变量名在赋值/取值时加不加符号都行：i% 和 i 指的是同一个变量
    MsgBox "i = " & i & "，l = " & l & "，arr(1) = " & arr%(1)

    ' ---- 不加后缀的坑：整数默认 Integer，装不下才升 Long ----
    ' 30000 * 2 会在 Integer 域里算 → 报"溢出"；加 & 强制按 Long 算就没事
    MsgBox "30000& * 2 = " & (30000& * 2)

    ' ---- 语句简写：一行写多条语句用冒号 : 分隔 ----
    Dim n As Long: n = 5
    Dim m As Long: m = 10
    MsgBox "n + m = " & (n + m)   ' 也可以写成 Dim n As Long: n = 5: MsgBox ...

    ' ---- 函数简写：带 $ 的版本返回 String，不带 $ 的返回 Variant ----
    Dim raw As String
    raw = "   ATE   "
    MsgBox "Trim  → [" & Trim(raw) & "]" & vbCrLf & _
           "Trim$ → [" & Trim$(raw) & "]"
    ' 文本处理优先用 Trim$ / Left$ / Mid$ / Right$ / UCase$ / LCase$ / Chr$
    ' 好处：返回 String 而不是 Variant，更快，也不会把 Null 传染给整个表达式

    ' ---- 更老的写法：DefType 语句（了解即可，加了 Option Explicit 就失去意义）----
    ' DefInt A-Z        ' 写在模块顶部（过程外面）：A~Z 未声明的变量默认当 Integer
    ' DefLng X-Z        ' 指定 X~Z 默认 Long；同类还有 DefBool / DefByte / DefCur
    '                   ' DefDate / DefDbl / DefSng / DefStr / DefObj / DefVar
    ' x = 5             ' 不 Dim 直接赋值 = 隐式声明为 Variant（现在的代码会直接报错）
End Sub
