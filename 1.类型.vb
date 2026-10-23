Option Explicit ' 强制"先声明后使用"，正式代码建议每个模块顶部都加上

' ============================================================
' 入口示例
' ============================================================
Sub HelloVBA()
    ' 单引号是注释，相当于 JS 的 //
    MsgBox "你好，VBA！我是未来的自动化专家。"

    ' 可以换行，也可以不换行，换行用 下划线加回车的形式
    ' 缩紧也可以不用缩紧，为了阅读缩紧

    ' 变量的几种类型（下面逐个演示，单独运行某个 Sub 也可以）
    ' DemoBasicTypes      ' 1.   基本数据类型
    ' DemoShortForm       ' 1.1  简写形式（类型声明字符）
    ' DemoVariantAndConst ' 2.   Variant、Const、类型转换
    ' DemoArray           ' 3.   数组
    ' DemoScope           ' 4.   作用域（进阶）
End Sub


' ============================================================
' 1. 基本数据类型：Dim 变量名 As 类型
'    声明成什么类型，就只能装什么（VBA 会自动做隐式转换）
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


' ============================================================
' 1.1 简写形式：类型声明字符（type declaration character）
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


' ============================================================
' 2. Variant：万能类型（不写 As 就是 Variant）
'    + Const 常量 + 常用类型转换函数
' ============================================================
Sub DemoVariantAndConst()
    Dim v                     ' 不写 As，默认就是 Variant，什么都能装
    v = 123
    MsgBox "现在是数字：" & (v + 1)
    v = "文本"                ' 同一变量换个类型照样用
    MsgBox "现在是文本：" & v & "，长度 " & Len(v)
    v = Array("VCC", "GND")    ' 还能装数组
    MsgBox "v 当前类型代码：" & VarType(v) & "（vbString=8, vbArray+Variant=8204）"

    ' 注意：Dim a, b As Integer 里只有 b 是 Integer，a 是 Variant
    '      要都声明就得写成 Dim a As Integer, b As Integer

    ' Const 常量：声明后不能改，适合放阈值、路径、测试参数
    Const MAX_RETRY As Long = 3
    Const VOLT_TOL As Double = 0.05
    ' MAX_RETRY = 5   ' ← 取消注释会报错：常量不能重新赋值

    ' ---- 常用类型转换 ----
    Dim t As String
    t = "22.5"

    ' 整数类：会取整，.5 时按银行家舍入取偶（CInt(22.5) = 22，不是 23）
    MsgBox "CInt: " & CInt(t)
    MsgBox "CLng: " & CLng(t)

    ' 浮点类：CDbl 是按"系统区域设置"去认小数点的
    '   本机实测 CDbl("22.5") 会报 运行时错误 6（溢出），连手写常量都炸，
    '   属于环境问题（区域设置里的"小数点符号"不是英文句点 .），不是代码问题
    ' MsgBox "CDbl: " & CDbl(t)      ' ← 这行在本机会炸，先注释掉

    ' 替代方案 1：Val —— 固定按英文 "." 解析，不看区域设置，永不抛异常
    MsgBox "Val: " & Val(t)

    ' 替代方案 2（推荐）：先校验再转换，脏数据不会静默变成 0
    '   ToDouble 是本文件末尾的辅助函数
    If IsNumeric(t) Then
        MsgBox "校验后转换: " & ToDouble(t)
    Else
        MsgBox "不是合法数字：[" & t & "]"
    End If

    MsgBox "CStr: " & CStr(123)
    MsgBox "CBool: " & CBool(1) & "（非 0 即 True）"
    MsgBox "IsNumeric(""abc""): " & IsNumeric("abc")

    ' 一句话记法：
    '   字符 → 整数  用 CInt / CLng（会取整）
    '   字符 → 浮点  别用 CDbl，用 Val 或 CDec（本机 CDbl 有坑）
    '   取整数部分  Fix（向 0 取整）/ Int（向下取整），负数时行为不同

    ' 判断变量到底是不是空的
    Dim x
    MsgBox "IsEmpty(x): " & IsEmpty(x)
End Sub


' ============================================================
' 3. 数组：Dim 名称(下标) As 类型
'    默认下界为 0（模块顶部写 Option Base 1 则改成 1）
'    ★ 坑：Dim a(3) 是 0~3 共 4 个元素，不是 3 个
'    ★ 和 JS 的区别：VBA 的「二维数组」是一个数组带两个下标（矩形），
'      不是 JS 那种数组套数组；不能写 g(1)(2)，也不能参差
' ============================================================
Sub DemoArray()
    ' ---- 3.1 一维固定数组 ----
    Dim fixedArr(1 To 3) As Integer   ' 明确写死 1~3，共 3 个
    Dim zeroBased(3) As Integer       ' 省略下界 → 0~3，共 4 个（新手最容易数错）

    fixedArr(1) = 10
    fixedArr(2) = 20
    fixedArr(3) = 30
    zeroBased(0) = 100
    zeroBased(3) = 400

    ' ---- 3.2 动态数组：用多少开多少 ----
    Dim dynArr() As String            ' 声明时不写长度（固定数组不能 ReDim）
    ReDim dynArr(1 To 2)              ' 运行期再定长度
    dynArr(1) = "Pass"
    dynArr(2) = "Fail"

    ReDim Preserve dynArr(1 To 3)     ' Preserve = 保留已有数据，只扩容
    dynArr(3) = "Retest"              ' 少了 Preserve 的话，前两个元素会被清空

    ' ---- 3.3 二维数组（行, 列）----
    Dim grid(1 To 2, 1 To 3) As Double ' 2×3 = 6 个格子，必须矩形，不能参差
    grid(1, 1) = 3.3                   ' 第 1 行第 1 列：引脚 1 的第 1 个电压值
    grid(2, 3) = 5#

    MsgBox "fixedArr(2) = " & fixedArr(2) & vbCrLf & _
           "dynArr(3) = " & dynArr(3) & vbCrLf & _
           "grid(1,1) = " & grid(1, 1) & vbCrLf & _
           "一维元素个数：" & (UBound(fixedArr) - LBound(fixedArr) + 1) & vbCrLf & _
           "行数：" & (UBound(grid, 1) - LBound(grid, 1) + 1) & _
           "，列数：" & (UBound(grid, 2) - LBound(grid, 2) + 1)
    ' UBound/LBound 的第二个参数是"第几维"，不写默认第 1 维

    ' 二维数组可以一次性刷进工作表，行/列天然对应（ATE 写报告常用）
    ' Range("A1").Resize(2, 3).Value = grid

    ' ---- 3.4 简写形式 ----
    ' ① 类型声明字符：Dim 名字 + 符号，数组照样能用
    Dim iArr%(1 To 3)                 ' = Dim iArr(1 To 3) As Integer
    Dim sArr$(1 To 2)                 ' = Dim sArr(1 To 2) As String
    Dim mArr#(1 To 2, 1 To 2)         ' = Dim mArr(1 To 2, 1 To 2) As Double
    Dim dyn$()                        ' 动态 String 数组的简写
    iArr%(1) = 5
    sArr$(1) = "VCC"

    ' ② Array()：一行塞一组值，省去逐个赋值
    Dim codes As Variant
    codes = Array("VCC", "GND", "CLK")
    MsgBox "第 3 个信号：" & codes(2) & _
           "，共 " & (UBound(codes) - LBound(codes) + 1) & " 个"

    ' ③ Split()：字符串直接拆成数组，比手写循环省事
    Dim pins() As String
    pins = Split("VCC,GND,CLK,RESET", ",")
    MsgBox "拆出 " & (UBound(pins) - LBound(pins) + 1) & " 个引脚，第 2 个是 " & pins(1)

    ' ④ Join()：数组拼回字符串（Split 的反操作）
    MsgBox "拼回去：" & Join(pins, " | ")

    ' ⑤ Filter()：按关键字过滤出子数组
    MsgBox "以 C 开头的：" & Join(Filter(pins, "C"), ",")

    ' ---- 3.5 常用操作 ----
    ' For Each 遍历一维数组（遍历变量必须是 Variant）
    Dim p As Variant, txt As String
    For Each p In codes
        txt = txt & p & " "
    Next
    MsgBox "遍历结果：" & txt

    ' 判断变量是不是数组、数组有没有分配过
    MsgBox "IsArray(codes) = " & IsArray(codes)

    ' Erase：动态数组变回"未分配"（之后 UBound 会报下标越界，必须重新 ReDim）
    Erase dynArr
    ReDim dynArr(1 To 1)
    dynArr(1) = "Fresh"
    MsgBox "重新 ReDim 后：" & dynArr(1)

    ' 注意：VBA 不支持数组整体赋值，arr2 = arr1 会编译报错，
    '      要循环逐个拷，或者两边都用 Variant 来装
End Sub


' ============================================================
' 4. 作用域：变量在哪一层声明，决定谁能用它
' ============================================================
Private gCounter As Long             ' 模块级：本模块所有 Sub 都能用

Sub DemoScope()
    Dim temp As Long                 ' 过程级：只在 DemoScope 里有效
    Static keep As Long              ' 过程级，但调用结束后值会保留

    temp = temp + 1                  ' 每次调用都从 0 开始，永远是 1
    keep = keep + 1                  ' 每次调用累加，用来做"调用次数"计数
    gCounter = gCounter + 1          ' 模块级，跨 Sub 共享

    MsgBox "temp = " & temp & vbCrLf & _
           "keep = " & keep & vbCrLf & _
           "gCounter = " & gCounter
End Sub


' ============================================================
' 工具函数：字符串 → Double（规避 CDbl 的区域设置坑）
'   1. 去首尾空格、去千分位逗号
'   2. IsNumeric 校验，脏数据抛错，而不是静默变成 0
'   3. 用 Val 解析，不受本机"小数点符号"设置影响
'   仪器回读、CSV 读入这类场景，统一走这个函数
' ============================================================
Function ToDouble(ByVal s As String) As Double
    Dim clean As String
    clean = Trim$(s)
    clean = Replace(clean, ",", "")          ' "1,234.5" → "1234.5"

    If Not IsNumeric(clean) Then
        Err.Raise vbObjectError + 1, "ToDouble", "不是合法数字：[" & s & "]"
    End If

    ToDouble = Val(clean)                    ' 刻意不用 CDbl
End Function
