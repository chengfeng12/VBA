Option Explicit

' ============================================================
' 1.4 数组：Dim 名称(下标) As 类型
'     默认下界为 0（模块顶部写 Option Base 1 则改成 1）
'     ★ 坑：Dim a(3) 是 0~3 共 4 个元素，不是 3 个
'     ★ 和 JS 的区别：VBA 的「二维数组」是一个数组带两个下标（矩形），
'       不是 JS 那种数组套数组；不能写 g(1)(2)，也不能参差
' ============================================================
Sub DemoArray()
    ' ---- 1.4.1 一维固定数组 ----
    Dim fixedArr(1 To 3) As Integer   ' 明确写死 1~3，共 3 个
    Dim zeroBased(3) As Integer       ' 省略下界 → 0~3，共 4 个（新手最容易数错）

    fixedArr(1) = 10
    fixedArr(2) = 20
    fixedArr(3) = 30
    zeroBased(0) = 100
    zeroBased(3) = 400

    ' ---- 1.4.2 动态数组：用多少开多少 ----
    Dim dynArr() As String            ' 声明时不写长度（固定数组不能 ReDim）
    ReDim dynArr(1 To 2)              ' 运行期再定长度
    dynArr(1) = "Pass"
    dynArr(2) = "Fail"

    ReDim Preserve dynArr(1 To 3)     ' Preserve = 保留已有数据，只扩容
    dynArr(3) = "Retest"              ' 少了 Preserve 的话，前两个元素会被清空

    ' ---- 1.4.3 二维数组（行, 列）----
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

    ' ---- 1.4.4 简写形式 ----
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

    ' ---- 1.4.5 常用操作 ----
    ' For Each 遍历一维数组（遍历变量必须是 Variant）
    Dim p As Variant, txt As String
    For Each p In codes
        txt = txt & p & " "
    Next
    MsgBox "遍历结果：" & txt

    ' 判断变量是不是数组
    MsgBox "IsArray(codes) = " & IsArray(codes)

    ' Erase：动态数组变回"未分配"（之后 UBound 会报下标越界，必须重新 ReDim）
    Erase dynArr
    ReDim dynArr(1 To 1)
    dynArr(1) = "Fresh"
    MsgBox "重新 ReDim 后：" & dynArr(1)

    ' 注意：VBA 不支持数组整体赋值，arr2 = arr1 会编译报错，
    '      要循环逐个拷，或者两边都用 Variant 来装
End Sub
