Option Explicit

' ============================================================
' 5.3 Dictionary 字典：查表 / 映射 / 统计的第一选择
'     为什么高频：pin → limit 映射、按 bin 分类计数、去重，
'     用数组要写两层循环，用字典一行搞定
' ------------------------------------------------------------
' 和 JS 对照：
'     const m = new Map()      →  Set d = CreateObject("Scripting.Dictionary")
'     m.set(k, v)              →  d(k) = v
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
' ============================================================
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
    ' d.Add "CLK", 1.2               ' 适合"键必须唯一，重复就是数据错"的场景

    ' ---- 删 ----
    d.Remove "GND"
    ' d.RemoveAll                    ' 全清
    MsgBox "删完后条目数：" & d.Count

    ' ---- 遍历：Keys / Items 都是 0 基数组 ----
    Dim k As Variant
    For Each k In d.Keys
        Debug.Print k & " → " & d(k)         ' 看立即窗口 Ctrl+G
    Next
    ' ★ 顺序不保证（虽然通常接近插入顺序），别拿它当"有序表"用
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
