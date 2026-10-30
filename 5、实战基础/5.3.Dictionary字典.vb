Option Explicit

' ============================================================
' 5.3 Dictionary 字典：查表 / 映射 / 统计的第一选择
'     为什么高频：pin → limit 映射、按 bin 分类计数、去重，
'     用数组要写两层循环，用字典一行搞定
' ------------------------------------------------------------
' 和 JS 对照：
'     const m = new Map()      →  Set d = CreateObject("Scripting.Dictionary")
'     m.set(k, v)              →  d(k) = v（新增或覆盖）
'                                 d.Add k, v（只新增，重复报 457）
'     m.get(k)                 →  d(k)     ← 键不存在时会"顺手新建"，见下面的坑
'     m.has(k)                 →  d.Exists(k)
'     m.size                   →  d.Count
'     [...m.keys()]            →  d.Keys()   返回 0 基数组
'     m.delete(k)              →  d.Remove k
'     m.clear()                →  d.RemoveAll
' ------------------------------------------------------------
' 两种建法：
'   后期绑定（推荐）：CreateObject("Scripting.Dictionary")
'                    不用勾引用、不限 Office 位数，拷到别的机器照样跑
'   早期绑定：工具→引用 勾 "Microsoft Scripting Runtime"，然后 Dim d As Dictionary
'            换来属性提示，代价是换机器可能得重新勾
'   两种写法的对比和 Dim d As New Dictionary 的坑见 DemoDictBinding
' ============================================================


' ============================================================
' 5.3.1 早期绑定 vs 后期绑定：为什么示例都写 Dim d As Object
'   后期绑定（本文件统一用）：
'     Dim d As Object
'     Set d = CreateObject("Scripting.Dictionary")
'   早期绑定（想写 Dim d As Dictionary 就得先做这一步）：
'     VBE → 工具 → 引用 → 勾 "Microsoft Scripting Runtime"
'     然后 Dim d As Dictionary / Dim d As New Dictionary
'   没勾引用就写 As Dictionary，编译期直接报「用户定义类型未定义」
' ============================================================
Sub DemoDictBinding()
    ' ---- 后期绑定：不依赖引用，拷到哪都能跑 ----
    Dim lateD As Object
    Set lateD = CreateObject("Scripting.Dictionary")
    lateD("VCC") = 3.3
    MsgBox "后期绑定：" & lateD("VCC")
    ' ★ 优点：不用勾引用；32/64 位 Office 都能跑
    ' ★ 代价：没有属性提示，成员名写错要到运行期才发现

    ' ---- 早期绑定：勾了引用才能编译 ----
    ' Dim earlyD As Dictionary
    ' Set earlyD = New Dictionary
    ' earlyD("VCC") = 3.3
    ' MsgBox "早期绑定：" & earlyD("VCC")
    ' ★ 优点：有 IntelliSense、成员名编译期检查
    ' ★ 代价：.xlsm 会记住这个引用，换机器/换 Office 位数时
    '   可能提示"找不到工程或库"，要重新勾
    ' ★ 写全名也行：Dim earlyD As Scripting.Dictionary

    ' ---- Dim d As New Dictionary 的坑 ----
    ' Dim autoD As New Dictionary
    ' If autoD Is Nothing Then ...     ' 这句本身就会创建对象
    ' ★ New 写在 Dim 里是"第一次用到才创建"，连 Is Nothing 也算用到，
    '   所以这个判断永远不成立；要显式控制就 Set autoD = New Dictionary

    ' ---- 两种可以混用，但没必要 ----
    ' Dim d As Dictionary
    ' Set d = CreateObject("Scripting.Dictionary")   ' 也合法
    ' 既然都勾引用了，直接用 New Dictionary 更清楚

    ' ★ Dictionary 不是 VBA 内置类型，是 scrrun.dll 里的 COM 对象；
    '   VBA 自带的只有 Collection（功能少，见 DemoDictTraps 第 5 条）
    ' ★ 引用路径：64 位 Office 是 System32\scrrun.dll，
    '   32 位 Office 是 SysWOW64\scrrun.dll；一般不用手填，勾名字即可
End Sub


Sub DemoDictBasics()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    ' ---- 增 / 改：同一个语法，键不存在就新增，存在就覆盖 ----
    d("VCC") = 3.3
    d("GND") = 0#
    d("VCC") = 3.31                 ' 覆盖，不报错
    MsgBox "条目数：" & d.Count      ' 2

    ' ---- 查：先 Exists 再取值 ----
    If d.Exists("VCC") Then MsgBox "VCC 的上限 = " & d("VCC")
    ' ★★ 最大的坑：d("不存在的键") 不报错，它会【新建】一个空条目，
    '    于是 Count 悄悄 +1，循环里读着读着数据就多了；
    '    JS 的 Map.get 返回 undefined，这里返回"空值 + 副作用"，务必先 Exists

    ' ---- Add：显式新增，键重复直接报错（457）----
    d.Add "CLK", 1.2                 ' 两个参数都必须给：Add key, item
    MsgBox "Add 后条目数：" & d.Count      ' 3
    ' d.Add "CLK", 1.5               ' ✘ 运行时错误 457：键已存在
    ' ★ d(k) = v 是「新增或覆盖」；d.Add 是「只新增」，重复直接报错
    ' ★ 适合"键必须唯一，重复就是数据错"的场景（引脚配置重复、批次号重复）
    ' ★ 重复键不确定时先 Exists 判断，或者 On Error 捕 457（见 3.3）

    If Not d.Exists("RESET") Then
        d.Add "RESET", 0#            ' 没有才 Add
    Else
        d("RESET") = 0#              ' 已有就走覆盖
    End If

    ' 想捕获 457 而不是让它弹错：
    ' On Error Resume Next
    ' d.Add "CLK", 1.5
    ' If Err.Number = 457 Then Debug.Print "CLK 重复，改为覆盖"
    ' Err.Clear
    ' On Error GoTo 0

    ' ---- 删 ----
    d.Remove "GND"
    ' d.RemoveAll                    ' 全清
    MsgBox "删完后条目数：" & d.Count      ' 3（VCC / CLK / RESET）

    ' ---- 遍历：Keys / Items 都是 0 基数组 ----
    Dim k As Variant
    For Each k In d.Keys
        Debug.Print k & " → " & d(k)         ' 看立即窗口 Ctrl+G
    Next
    ' ★ 顺序不保证（虽然通常接近插入顺序），别拿它当"有序表"用
End Sub


' ============================================================
' 5.3.2 修改键：没有 Rename 方法，用 Key 属性
'   d.Key(旧键) = 新键        ' 改键，值跟着走
'   ★ 旧键不存在 → 悄悄新建一个键，值是 Empty（不报错！）
'   ★ 新键已存在 → 运行时错误 457（重复键）
'   ★ 想稳一点：先 Exists 两头都判，或者整体重建
' ============================================================
Sub DemoDictChangeKey()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("VCC") = 3.3
    d("GND") = 0#

    ' ---- ① 直接改：值不变，Count 不变 ----
    d.Key("VCC") = "VCC_3V3"
    MsgBox "改后 Count：" & d.Count & _
           "，VCC 还在吗：" & d.Exists("VCC") & _
           "，新键的值：" & d("VCC_3V3")
    ' 输出：2 / False / 3.3
    ' ★ 别写成 d("VCC") = "VCC_3V3" —— 那是改「值」，不是改键；
    '   改键只能 d.Key(旧键) = 新键

    ' ---- ② 新键已存在 → 457 ----
    ' d.Key("GND") = "VCC_3V3"        ' ✘ 运行时错误 457：重复键

    ' ---- ③ 旧键不存在 → 悄悄新建一个空值键（最容易踩）----
    ' d.Key("NC") = "NC_1"            ' 不报错，Count +1，d("NC_1") = Empty
    ' ★ 和 d("不存在的键") 一样，是"读写都能顺手建键"的坑；
    '   改名前一定要先 d.Exists(旧键)

    ' ---- ④ 稳妥写法：两头都判 ----
    Dim oldKey As String, newKey As String
    oldKey = "GND"
    newKey = "GND_0V"
    If Not d.Exists(oldKey) Then
        MsgBox "旧键不存在：" & oldKey
    ElseIf d.Exists(newKey) Then
        MsgBox "新键已存在，别硬改：" & newKey
    Else
        d.Key(oldKey) = newKey
        MsgBox "改名成功：" & newKey & " = " & d(newKey)   ' GND_0V = 0
    End If

    ' ---- ⑤ 批量改名 / 规范化：重建更安全 ----
    ' 一个个改容易撞 457，推荐新建一个字典再灌一遍
    Dim src As Object, dst As Object, k As Variant
    Set src = CreateObject("Scripting.Dictionary")
    src(" vcc ") = 3.3
    src("gnd") = 0#

    Set dst = CreateObject("Scripting.Dictionary")
    For Each k In src.Keys
        dst(UCase$(Trim$(CStr(k)))) = src(k)      ' 新键冲突时后值覆盖
    Next
    MsgBox "规范化后：" & dst.Count & " 个，VCC=" & dst("VCC") & "，GND=" & dst("GND")

    ' ★ 想"冲突就报错"而不是悄悄覆盖：
    '   赋值前先 If dst.Exists(规范化后的键) Then Err.Raise ...
    ' ★ Remove + Add 是"删了再加"，通常会挪到末尾；d.Key 直接改不一定，
    '   反正顺序本来就不保证，需要稳定顺序就另存一个键数组
End Sub


' ============================================================
' ATE 用例 1：引脚名 → 上下限的映射表
'   离线 limit 表读进来后，就长这样；取值时按引脚名查，不用记列号
' ============================================================
Sub DemoPinLimit()
    Dim limits As Object
    Set limits = CreateObject("Scripting.Dictionary")

    limits("VCC") = Array(3.3, 3.6)      ' 值本身可以是数组
    limits("GND") = Array(-0.1, 0.1)
    limits("CLK") = Array(1.5, 2#)

    Dim pins As Variant, i As Long, pin As String
    pins = Array("VCC", "GND", "CLK", "VIO")
    For i = LBound(pins) To UBound(pins)
        pin = CStr(pins(i))
        If limits.Exists(pin) Then
            Debug.Print pin & " 限值 " & limits(pin)(0) & " ~ " & limits(pin)(1)
        Else
            Debug.Print pin & " 没有配置限值 → 跳过或按配置缺失处理"
            ' ★ 未配置的引脚一定要显式处理：默认放行会造成"漏测还判 Pass"
        End If
    Next
End Sub


' ============================================================
' ATE 用例 2：按结果 / bin 分类计数（良率统计的最小单元）
' ============================================================
Sub DemoCountByBin()
    Dim datas As Variant, i As Long, key As String
    Dim cnt As Object
    Set cnt = CreateObject("Scripting.Dictionary")

    datas = Array("Pass", "Pass", "Fail", "Pass", "Bin5", "Fail")
    For i = LBound(datas) To UBound(datas)
        key = CStr(datas(i))
        If cnt.Exists(key) Then
            cnt(key) = cnt(key) + 1          ' 已有：+1
        Else
            cnt(key) = 1                     ' 没有：初始化为 1
        End If
    Next

    Dim k As Variant, total As Long
    For Each k In cnt.Keys
        total = total + cnt(k)
        Debug.Print k & " : " & cnt(k)
    Next
    Debug.Print "合计 " & total & " 颗"
End Sub


' ============================================================
' 几个必须知道的细节
' ============================================================
Sub DemoDictTraps()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    ' ---- 1) 键默认区分大小写；忽略大小写要在"加键之前"设 CompareMode ----
    d.CompareMode = 1                ' 1 = vbTextCompare，0 = vbBinaryCompare（默认）
    d("vcc") = 3.3
    MsgBox "大小写不敏感？" & d.Exists("VCC")        ' True
    ' ★ CompareMode 一旦有键就不能再改，会报错 5

    ' ---- 2) 遍历时不能直接 Remove ----
    Dim k As Variant, keysLeft As Variant
    keysLeft = d.Keys                ' 先把键拷出来
    For Each k In keysLeft
        ' d.Remove k                   ' 直接删会跳项 / 报错，所以先拷
    Next
    ' ★ 正确姿势：keysLeft = d.Keys → 再循环；RemoveAll 是另一回事，可以随时用

    ' ---- 3) 键只能是字符串或数值 ----
    ' d(Array(1, 2)) = "x"           ' 报错：不能拿数组当键
    ' 需要"复合键"时自己拼：d("VCC|3.3|25C") = 结果

    ' ---- 4) 值要放对象，赋值也要 Set ----
    ' Dim rng As Object
    ' Set rng = ThisWorkbook.Sheets(1).Range("A1")
    ' Set d("A1") = rng              ' ✓ 对象值用 Set
    ' d("A1") = rng                  ' ✗ 会走默认属性，存进去的是值不是对象

    ' ---- 5) 与 Collection 的区别 ----
    ' Collection：有 Add/Remove/Item，能按键取，但【不能判断键是否存在】，
    '             也不方便遍历键；要判重就得靠 On Error
    ' Dictionary：有 Exists / Count / Keys / Items，日常一律用它
End Sub
