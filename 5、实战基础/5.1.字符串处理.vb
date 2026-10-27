Option Explicit

' ============================================================
' 5.1 字符串处理：引脚名、版本号、日志、结果拼接天天要用
' ------------------------------------------------------------
' 本文件按「想干什么」分组，忘了函数名就按组找：
'   去空白    Trim / LTrim / RTrim（只去 Chr(32)，全角空格和 Tab 不管）
'   查找      InStr / InStrRev / StrComp / = 比较
'   截取      Left / Mid / Right / Mid 语句（原地改写）
'   拆合      Split（limit、多分隔符）/ Join / Filter
'   替换      Replace（起点、次数、比较模式）
'   大小写    UCase / LCase / StrConv（ProperCase、全半角）
'   字符码    Asc / AscW / Chr / ChrW / StrReverse / String$ / Space$
'   模式      Like（? * # [] [!]）
'   数值互转  CStr / Str / Val / Format / IsNumeric
'   拼接      数组 + Join（避免循环 & 的 O(n²)）
'   正则      VBScript.RegExp（Like 不够用时）
'   工具函数  StartsWith / EndsWith / PadLeft / PadRight / TrimChars / SplitAny
' ------------------------------------------------------------
' 和 JS 对照（几乎一一对应，只有下标规则不同）：
'     s.trim()              → Trim$(s)              LTrim$ / RTrim$ 只去一边
'     s.indexOf("x")        → InStr(s, "x")        1 基，找不到返回 0
'     s.lastIndexOf("x")    → InStrRev(s, "x")     注意参数顺序和 InStr 相反
'     s.includes("x")       → InStr(s, "x") > 0
'     s.startsWith("x")     → Left$(s, Len(x)) = x 无内置，见文件末尾
'     s.endsWith("x")       → Right$(s, Len(x)) = x
'     s.slice(1, 4)         → Mid$(s, 2, 3)        第二参数是「起点」
'     s.slice(0, 3)         → Left$(s, 3)
'     s.slice(-3)           → Right$(s, 3)
'     s.at(-1)              → Right$(s, 1)
'     s.split(",")          → Split(s, ",")        得到 0 基数组
'     arr.join(",")         → Join(arr, ",")
'     s.replaceAll(a, b)    → Replace(s, a, b)     默认全替换，只做字面量
'     s.replace(/re/, x)    → CreateObject("VBScript.RegExp").Replace
'     s.match(/re/)         → RegExp.Execute
'     s.toUpperCase()       → UCase$(s)
'     s.repeat(3)           → String$(3, s)（只适用单字符，多字符要循环拼）
'     s.padStart(4, "0")    → 自写 PadLeft（数值用 Format$）
'     s.charCodeAt(0)       → AscW(Left$(s, 1))
'     String.fromCharCode() → ChrW(n)
'     s1.localeCompare(s2)  → StrComp(s1, s2, vbTextCompare)
'     /^VCC\d+$/            → s Like "VCC##"（简单）或 RegExp（复杂）
'     s1 === s2             → StrComp(s1, s2, vbBinaryCompare) = 0
' ------------------------------------------------------------
' 三条贯穿全文的规则：
'   1. 文本函数有带 $ 和不带 $ 两版：
'        Trim$(s) → 返回 String，快；参数是 Null 时直接报错 94
'        Trim(s)  → 返回 Variant；参数是 Null 时返回 Null
'      数据可能为 Null 时用不带 $ 的版本并自己判 IsNull；
'      确定是 String 就用 $ 版，少一层 Variant 包装
'   2. 所有「位置」都是 1 基：InStr 的返回值、InStr 的起点、
'      Mid/Left/Right 的参数、Replace 的 start；
'      只有 Split / Filter 拆出来的数组是 0 基 —— 混用必错位
'   3. VBA 没有 startsWith / endsWith / padStart / padEnd / replaceAll /
'      按正则 split，用本文件末尾的工具函数组合，别每次手写
' ------------------------------------------------------------
' 顺带记一下几个常量：
'   vbCrLf = Chr(13) & Chr(10)   vbCr = Chr(13)   vbLf = Chr(10)   vbTab = Chr(9)
'   vbNullChar = Chr(0)          vbNullString = ""（不占内存的空串常量）
' ============================================================


' ============================================================
' 一、去空白：Trim 只认普通空格
' ============================================================
Sub DemoTrim()
    Dim s As String, t As String
    s = "  VCC_3V3_1.25V  "

    MsgBox "Trim ：[" & Trim$(s) & "]"        ' [VCC_3V3_1.25V]
    MsgBox "LTrim：[" & LTrim$(s) & "]"       ' [VCC_3V3_1.25V  ]
    MsgBox "RTrim：[" & RTrim$(s) & "]"       ' [  VCC_3V3_1.25V]
    ' ★ 三个都只动首尾，不动中间；中间连续空格压成一个要自己循环或上正则

    ' ---- 只认普通空格 Chr(32)：下面这些「看不见的空白」一个都不去 ----
    t = ChrW(12288) & "VCC_3V3" & vbTab       ' 全角空格 + Tab
    MsgBox "原长 " & Len(t) & "，Trim 后长 " & Len(Trim$(t))    ' 9 / 9
    ' ★ 从网页、别的系统导出的表里常混全角空格（ChrW(12288)）、
    '   不换行空格（ChrW(160)）、Tab：肉眼和普通空格一模一样，
    '   但 Trim / InStr / 字典键全都不认 → 看着「一样的值」，比较却不相等

    ' 统一清洗：把各种空白都换成普通空格，再 Trim
    t = Replace(t, ChrW(12288), " ")
    t = Replace(t, ChrW(160), " ")
    t = Replace(t, vbTab, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    MsgBox "清洗后：[" & Trim$(t) & "]"        ' [VCC_3V3]

    ' ---- 判空必须先 Trim ----
    MsgBox "全是空格算空吗：" & (Len(Trim$("     ")) = 0)     ' True
    ' ★ 用户输入 "   " 时 Len(s) = 3，直接 If Len(s) = 0 会当成有值

    ' ---- BOM：读 UTF-8 文件（5.2 的 ReadUtf8）时首字符可能是它 ----
    Dim line As String
    line = ChrW(&HFEFF&) & "VCC,3.31,Pass"
    If Left$(line, 1) = ChrW(&HFEFF&) Then line = Mid$(line, 2)
    MsgBox "去 BOM 后：" & Left$(line, 3)     ' VCC
    ' ★ BOM 是非打印字符，Debug.Print 里也看不见；
    '   「文件第一行的引脚名死活匹配不上」多半就是它
    ' ★ 十六进制字面量不加 & 时按 16 位整数解释：&HFEFF 会变成负数，
    '   写 ChrW(&HFEFF&) 才是 65279（见 DemoCharCode 的 &HFFFD&）

    ' ---- 去自定义字符（引号、括号）：Trim 做不到，用末尾的 TrimChars ----
    MsgBox "去引号：" & TrimChars(Chr(34) & "VCC" & Chr(34), Chr(34))   ' VCC
    MsgBox "去括号：" & TrimChars("(3.31)", "()")                       ' 3.31
    ' ★ TrimChars 只去首尾出现的、chars 里列出的字符，中间不动（和 Trim 同语义）

    ' ★ Trim(Null) 返回 Null；Trim$(Null) 报错 94（无效的 Null 使用）。
    '   从单元格读出来的值先 v & "" 兜底，再交给 $ 版函数
End Sub


' ============================================================
' 二、查找与比较：找不到返回 0，不是 -1
' ============================================================
Sub DemoSearch()
    Dim s As String
    s = "  VCC_3V3_1.25V  "

    ' ---- InStr([起点,] 原串, 要找的 [, 比较模式]) ----
    MsgBox "InStr 找 V3：" & InStr(s, "V3")          ' 8（1 基）
    MsgBox "找不到 GND：" & InStr(s, "GND")          ' 0，不是 -1
    MsgBox "从第 9 位往后找 V：" & InStr(9, s, "V")   ' 15
    MsgBox "起点超过长度：" & InStr(100, s, "V")      ' 0，不报错
    ' ★ 判存在一律写 If InStr(s, "GND") > 0 Then，
    '   别写 If InStr(s, "GND") Then（0 是 False，但 > 0 更直白）

    ' 忽略大小写：第 4 个参数传 vbTextCompare
    MsgBox "忽略大小写找 vcc：" & InStr(1, s, "vcc", vbTextCompare)   ' 3
    ' ★ 参数顺序是 (起点, 原串, 要找的, 比较模式)，起点不能省掉直接写比较模式
    ' ★ 空针返回起点：InStr(5, s, "") = 5；原串为空时返回 0

    ' ---- InStrRev：从右往左找，参数顺序和 InStr 相反 ----
    ' InStrRev(原串, 要找的 [, 从第几位往前找] [, 比较模式])
    MsgBox "最后一个下划线：" & InStrRev(s, "_")           ' 10
    MsgBox "从第 9 位往前找：" & InStrRev(s, "_", 9)       ' 6
    MsgBox "忽略大小写找 v：" & InStrRev(s, "v", -1, vbTextCompare)   ' 15
    ' ★ 省略 start 等价于 -1（从末尾开始）；返回值仍是「从左数」的位置
    ' ★ 只换最后一次出现就靠它定位（见 DemoReplace 末尾）

    ' ---- StrComp：三向比较，返回 -1 / 0 / 1 ----
    MsgBox "区分大小写：" & StrComp("VCC", "vcc")                     ' -1
    MsgBox "忽略大小写：" & StrComp("VCC", "vcc", vbTextCompare)      ' 0
    MsgBox "字典序：" & StrComp("abc", "abd")                         ' -1
    If StrComp("VCC", "vcc", vbTextCompare) = 0 Then MsgBox "同一个引脚"
    ' ★ 返回的是符号不是差值，别拿它做减法
    ' ★ 任一参数是 Null 时返回 Null，不是 0/-1；数据来自单元格时先 v & ""

    ' ---- = 比较：默认区分大小写 ----
    MsgBox "VCC = vcc：" & ("VCC" = "vcc")        ' False
    ' ★ 想忽略大小写别先 UCase 两边（多两次内存分配），
    '   用 StrComp(..., vbTextCompare) = 0 一步到位，意图也更清楚

    ' ---- Option Compare：模块级开关，本仓库故意不用 ----
    '   Option Compare Binary（默认）→ = < > Like、省略 compare 的 InStr 全区分大小写
    '   Option Compare Text           → 同一模块内全不区分大小写
    '   它只能写在模块最顶部，一改影响整个模块（连 If 判断、Select Case 都变），
    '   排查时极难发现，所以需要忽略大小写时一律显式传 vbTextCompare

    ' ---- 统计出现次数：起点要 +1 往前挪，否则永远停在第一个 ----
    Dim p As Long, n As Long
    p = 1
    Do
        p = InStr(p, "VCC_3V3_VCC_5V", "VCC")
        If p = 0 Then Exit Do
        n = n + 1
        p = p + 1                 ' 找不重叠的可以 p + Len("VCC")
    Loop
    MsgBox "VCC 出现 " & n & " 次"        ' 2
End Sub


' ============================================================
' 三、截取：Left / Mid / Right，越界静默返回空串
' ============================================================
Sub DemoSlice()
    Dim t As String
    t = Trim$("  VCC_3V3_1.25V  ")        ' "VCC_3V3_1.25V"，长 13

    MsgBox "Left 3：" & Left$(t, 3)        ' VCC
    MsgBox "Mid 从 5 取 3：" & Mid$(t, 5, 3)   ' 3V3
    MsgBox "Right 3：" & Right$(t, 3)      ' 25V
    MsgBox "Mid 省略长度：" & Mid$(t, 5)   ' 3V3_1.25V（一直取到结尾）

    ' ---- 边界行为：不报错，但结果容易出乎意料 ----
    MsgBox "取 0 个：[" & Left$(t, 0) & "]"          ' []
    MsgBox "超长：[" & Left$(t, 100) & "]"           ' 整串，不报错
    MsgBox "起点越界：[" & Mid$(t, 100, 5) & "]"     ' []，也不报错
    ' ★ 越界静默返回空串，比 JS 的 undefined 更隐蔽 —— 取完记得判 Len
    ' ★ Left$(t, -1) 和 Mid$(t, 0, 2) 会报运行时错误 5
    '   （Mid 的起点必须 >= 1，Left/Right 的长度必须 >= 0）
    ' ★ Mid 的第二个参数是「起点」不是「结尾下标」，这是从 JS slice 转过来
    '   最容易写错的地方：slice(1,4) → Mid(s, 2, 3)

    ' ---- Mid 语句：原地改写一个已存在的字符串变量 ----
    Dim u As String
    u = "VCC-3V3"
    Mid(u, 4, 3) = "GND"                  ' 把第 4 位起 3 个字符换成 "GND"
    MsgBox u                              ' VCC-GND

    u = "VCC-3V3"
    Mid(u, 4, 3) = "5V"                   ' 只覆盖前 2 个字符，剩下那个不动
    MsgBox u                              ' VCC-5V3
    ' ★ Mid 语句不会改变原串长度：短了补不上、长了被截断；
    '   想改长度只能整体重新赋值：u = Left$(u, 3) & "GND"
    ' ★ 左边必须是「能赋值的变量」，字符串字面量改不了；
    '   起点 < 1 报错 5。日常更推荐「用函数拼新串」，可读性更好

    ' ---- 实战两件套：去前缀 / 去后缀（VBA 没有 startsWith 的替代写法）----
    u = "VCC_3V3"
    If Left$(u, 4) = "VCC_" Then u = Mid$(u, 5)
    MsgBox u                              ' 3V3

    u = "3.31V"
    If Right$(u, 1) = "V" Then u = Left$(u, Len(u) - 1)
    MsgBox u                              ' 3.31

    MsgBox "最后一个字符：" & Right$(t, 1)     ' V（等价 JS 的 s.at(-1)）

    ' ---- 封成函数更不容易漏 Len ----
    MsgBox "以 VCC 开头：" & StartsWith(t, "VCC")    ' True
    MsgBox "以 5V 结尾：" & EndsWith(t, "5V")        ' True
    ' ★ 判断前缀也可以写 t Like "VCC*"，但每次都要自己记 * 的位置；
    '   函数版对「空前缀」「前缀比原串长」这类边界已经处理好了
End Sub


' ============================================================
' 四、拆分与合并：Split 的数组是 0 基
' ============================================================
Sub DemoSplitJoin()
    Dim parts As Variant, s As String, i As Long

    parts = Split("VCC,3.31,Pass", ",")
    MsgBox "段数：" & (UBound(parts) - LBound(parts) + 1)     ' 3
    MsgBox "第 1 段：" & parts(0)                             ' VCC，0 基！

    ' ---- 分隔符省略时默认是空格 ----
    parts = Split("VCC 3.31 Pass")
    MsgBox "按空格拆：" & parts(1)                            ' 3.31

    ' ---- 空串和连续分隔符：会拆出空元素，不会自动合并 ----
    parts = Split("", ",")
    MsgBox "空串：UBound = " & UBound(parts) & "，内容 [" & parts(0) & "]"
    ' ★ 空串得到的是「1 个空元素」的数组，UBound = 0 不是 -1，
    '   所以 parts(0) 能取到空串，别把它当成有效值

    parts = Split("VCC,,GND,", ",")
    MsgBox "连续/结尾分隔符：" & (UBound(parts) + 1) & " 段"   ' 4，其中 2 个空串
    ' ★ 脏数据里最常见的就是这种空段，取完统一 Trim 并跳过空串

    ' ---- limit：最多拆几段，剩下的整段塞在最后 ----
    parts = Split("2026-10-26,09:30:00,VCC,Pass", ",", 3)
    MsgBox "第 3 段：" & parts(2)                             ' VCC,Pass
    ' 签名：Split(原串 [, 分隔符] [, limit] [, 比较模式])
    '   limit = -1（默认）全部拆；= n 只拆成 n 段

    ' ---- 分隔符可以是多字符，但只认「一个」分隔符 ----
    parts = Split("VCC | GND | CLK", " | ")
    MsgBox "多字符分隔符：" & Join(parts, "/")                 ' VCC/GND/CLK

    ' 多个不同分隔符（逗号、分号、Tab）：先统一成一种再拆
    s = "VCC,GND;CLK" & vbTab & "RESET"
    s = Replace(s, ";", ",")
    s = Replace(s, vbTab, ",")
    MsgBox "统一后：" & Join(Split(s, ","), " / ")
    ' ★ Split 不支持「多个分隔符」，也不支持正则；
    '   或者直接用末尾的 SplitAny(s, ",;" & vbTab)

    ' ---- Filter：按关键字筛子数组（包含匹配，不是前缀匹配）----
    Dim pins As Variant
    pins = Split("VCC,GND,CLK,RESET,VCC_3V3", ",")
    MsgBox "含 VCC：" & Join(Filter(pins, "VCC"), ",")              ' VCC,VCC_3V3
    MsgBox "不含 VCC：" & Join(Filter(pins, "VCC", False), ",")     ' GND,CLK,RESET
    MsgBox "忽略大小写：" & Join(Filter(pins, "vcc", True, vbTextCompare), ",")
    ' 签名：Filter(数组, 关键字 [, 是否保留匹配] [, 比较模式])
    ' ★ 默认区分大小写；返回的数组永远 0 基
    ' ★ 一个都没匹配上时返回空数组，直接 UBound 会报「下标越界 9」，
    '   先 Join 出来看一眼，或自己写循环数个数

    ' ---- Join：数组拼回字符串 ----
    MsgBox "合回去：" & Join(pins, " | ")
    ' ★ 只能一维数组；元素不是字符串会隐式转换（数值按区域设置的小数点）
    ' ★ 拼多行文本用 Join(lines, vbCrLf)，比循环 & 快一个数量级（见 DemoConcat）

    ' ---- 拆完通常要逐段 Trim ----
    parts = Split("VCC, GND , CLK", ",")
    For i = LBound(parts) To UBound(parts)
        parts(i) = Trim$(parts(i))
    Next
    MsgBox "去空格后：[" & Join(parts, "][") & "]"

    ' ---- 也能直接赋给 String 数组 ----
    Dim sa() As String
    sa = Split("VCC,GND", ",")
    MsgBox "String 数组第 2 个：" & sa(1)     ' GND
End Sub


' ============================================================
' 五、替换：默认全替换、区分大小写、纯字面量
' ============================================================
Sub DemoReplace()
    ' 签名：Replace(原串, 找什么, 换成什么 [, 起始位置] [, 次数] [, 比较模式])
    '   起始位置 1 基；次数 -1 = 全部；找不到就原样返回，不报错
    MsgBox Replace("VCC-vcc", "vcc", "GND")                          ' VCC-GND
    MsgBox Replace("VCC-vcc", "vcc", "GND", 1, -1, vbTextCompare)    ' GND-GND
    MsgBox Replace("a-a-a", "a", "b", 1, 2)                          ' b-b-a（只换 2 个）
    MsgBox Replace("a-a-a", "a", "b", 3)                             ' a-b-b（从第 3 位起）

    ' vbBinaryCompare(0) 区分大小写 / vbTextCompare(1) 不区分
    ' ★ 这两个常量在 StrComp、InStr、Split、Filter 里都在用，值就是 0 和 1，
    '   但写名字比写数字清楚得多

    ' ---- 删字符 = 换成空串 ----
    MsgBox Replace("VCC-3.3V", "-", "")          ' VCC3.3V
    ' ★ 别顺手拿它删小数点：Replace("3.310", ".", "") = "3310"，
    '   数值会被放大 1000 倍

    ' ---- 双引号：字面量里写两遍，或者用 Chr(34) ----
    Dim q As String
    q = Chr(34)
    MsgBox "一个字面引号：[" & q & "] 等价于 [" & """" & "]"
    MsgBox "去掉首尾引号：" & Replace(q & "3,310" & q, q, "")     ' 3,310
    ' ★ 拼接 CSV 时用 Chr(34) 比数一串 """" 更不容易错

    ' ---- 统一换行：Windows 是 CRLF，Unix 是 LF ----
    Dim text As String
    text = "a" & vbCrLf & "b" & vbCrLf & "c"
    text = Replace(text, vbCrLf, vbLf)           ' 先换 CRLF
    text = Replace(text, vbCr, vbLf)             ' 再处理孤立的 CR
    MsgBox "行数：" & (UBound(Split(text, vbLf)) + 1)     ' 3
    ' ★ 顺序反了会剩下孤立的 LF：先换 CRLF 再换单个 CR

    ' ---- 只替换最后一次出现：Replace 做不到，InStrRev 定位后拼 ----
    Dim u As String, p As Long
    u = "VCC_3V3_VCC"
    p = InStrRev(u, "VCC")
    If p > 0 Then u = Left$(u, p - 1) & "GND" & Mid$(u, p + 3)
    MsgBox u                                     ' VCC_3V3_GND

    ' ---- Replace 是纯字面量，不是正则 ----
    ' ★ 把「连续多个空格压成一个」「去掉所有括号及里面的内容」这类需求
    '   交给 DemoRegex，别用 Replace 硬凑
    ' ★ 循环里反复 Replace 每次都新建字符串，几十万行时很明显；
    '   能一次 Replace 完就别拆成多次
End Sub


' ============================================================
' 六、大小写与全半角：比较时别先 UCase
' ============================================================
Sub DemoCase()
    MsgBox UCase$("pass") & " / " & LCase$("FAIL")     ' PASS / fail
    MsgBox "中文不受影响：" & UCase$("引脚电压")        ' 引脚电压

    ' ---- ProperCase：每个单词首字母大写（只对英文有效）----
    MsgBox StrConv("vcc pin", vbProperCase)            ' Vcc Pin
    ' ★ 引脚名 / 信号名一般不这么做（VCC 会变成 Vcc），
    '   适合人名、表头这类展示文本

    ' ---- 做键 / 分组前统一大小写 ----
    Dim key As String
    key = UCase$(Trim$(" vcc_3v3 "))
    MsgBox key                                         ' VCC_3V3
    ' ★ 拿它当字典键（5.3）能避免 "VCC" 和 "vcc" 被算成两组

    ' ---- 全角 ↔ 半角：处理从 Excel / 网页 / 日文系统导出的数据 ----
    MsgBox StrConv("VCC123", vbWide)                   ' ＶＣＣ１２３
    MsgBox StrConv("ＶＣＣ１２３", vbNarrow)           ' VCC123
    ' ★ 只转「有宽度变体」的字符（ASCII、片假名），汉字不变
    ' ★ 依赖系统区域设置，非中日文环境可能无效

    ' ---- 别拿 StrConv 当编码转换 ----
    ' StrConv(s, vbUnicode) / vbFromUnicode 是按系统代码页把每个字节
    ' 拉宽 / 压窄，不是「GBK ↔ UTF-8」这种换码；
    ' 文件编码转换一律用 5.2 的 ADODB.Stream

    ' ★ 比较时别先 UCase 再比：
    '     UCase$(a) = UCase$(b)            多分配两个临时字符串
    '     StrComp(a, b, vbTextCompare) = 0 一步到位，意图明确
End Sub


' ============================================================
' 七、字符码与单字符处理：Asc / Chr / StrReverse
' ============================================================
Sub DemoCharCode()
    MsgBox "Asc(""A"")：" & Asc("A")               ' 65
    MsgBox "Chr(65)：" & Chr$(65)                  ' A
    MsgBox "Chr(9) 是 Tab，Len = " & Len(Chr(9))   ' 1

    ' ---- 中文 / Unicode：用 W 版，别用 Asc 猜 ----
    MsgBox "AscW(""中"")：" & AscW("中")           ' 20013
    MsgBox "ChrW(&H4E2D&)：" & ChrW(&H4E2D&)       ' 中
    ' ★ Asc 返回「当前代码页」的编码，中文在 GBK 下可能是负数，
    '   跟区域设置绑定；判断字符一律用 AscW
    ' ★ AscW 是 16 位有符号数，高位为 1 的码点会变负，统一写成下面这样：
    MsgBox "码点：" & (CLng(AscW(ChrW(&HFFFD&))) And &HFFFF&)    ' 65533
    ' ★ 十六进制字面量不加 & 时按 16 位解释：&HFFFD 是 -3，&HFFFD& 才是 65533；
    '   写 ChrW 的参数时养成加 & 的习惯，超过 &H7FFF 的码点才不会出错

    ' ---- Chr(0)：能存住，但一交给 API 就被当成字符串结束 ----
    MsgBox "Chr(0) 的 Len = " & Len(Chr(0))        ' 1
    ' ★ 写进某些控件 / Win32 API 时后面的内容会悄悄丢掉，
    '   清洗脏数据时优先把它 Replace 成空串

    ' ---- 反转 ----
    MsgBox "StrReverse：" & StrReverse("VCC")      ' CCV

    ' ---- 构造字符串：String$ 重复字符，Space$ 造空格 ----
    MsgBox "重复字符：" & String$(3, "0")           ' 000
    MsgBox "传多字符只取第一个：" & String$(3, "ab") ' aaa
    MsgBox "n 个空格：[" & Space$(5) & "]"
    ' ★ 长度给负数会报错 5；String$ 的第二个参数是 Variant，
    '   给数字时按 Chr(数字) 处理：String$(3, 48) 也是 "000"

    ' ---- 逐字符扫描：日志清洗、挑数字、数中文时用 ----
    Dim s As String, c As String, i As Long, digits As Long
    s = "VCC-3V3"
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c Like "#" Then digits = digits + 1
    Next
    MsgBox "数字字符个数：" & digits                ' 2
    ' ★ 逐字符用 Mid$(s, i, 1)，Len 是字符数；
    '   B 系列函数（MidB / LenB / AscB）按字节算，中文会被切成半个字，
    '   只在处理二进制或编码时才碰

    ' ---- Len 是字符数，LenB 是 UTF-16 字节数（每个字符固定 2 字节）----
    MsgBox "Len：" & Len("VCC电压") & "，LenB：" & LenB("VCC电压")   ' 5 / 10
    ' ★ VBA 字符串内部是 UTF-16，LenB = 字符数 × 2，
    '   英文也是 2 字节/字符 —— 所以不能拿「LenB - Len」去数中文
    ' ★ 要按存储/接口的编码算字节（如 GBK 下「电压」占 4 字节），
    '   用 LenB(StrConv(s, vbFromUnicode))：VB 官方文档里的 LenMbcs 写法
    '   LenB(StrConv("VCC电压", vbFromUnicode)) 在中文 GBK 环境下 = 7
    '   （换成 UTF-8 代码页会变 9，所以这个值只能当「当前环境的实测结果」）
    ' ★ 校验「字段是否超长」时先想清楚是按字符还是按字节，
    '   数据库 VARCHAR(n) 的 n 到底怎么算，得看具体库（见 5.2 的 ANSI/UTF-8）

    ' ★ Emoji / 生僻字在 VBA 里是「代理对」，Len 算 2，
    '   Mid$(s, i, 1) 只拿到半个字符；引脚名里一般没有，日志里可能有
End Sub


' ============================================================
' 八、模式匹配 Like：正则的轻量替代
' ============================================================
Sub DemoLike()
    '   ?  任意单个字符        *  任意多个字符（含 0 个）
    '   #  任意单个数字 0-9    [abc] 列表里任一个    [A-Z] 区间
    '   [!0-9] 不在列表里      [] 空串
    MsgBox "VCC12 像 VCC##：" & ("VCC12" Like "VCC##")     ' True
    MsgBox "VCC1A 像 VCC##：" & ("VCC1A" Like "VCC##")     ' False
    MsgBox "VCC 像 VCC*：" & ("VCC" Like "VCC*")           ' True
    MsgBox "VCC1x 像 VCC#"：" & ("VCC1x" Like "VCC#")      ' False
    ' ★ Like 天然是「整串匹配」，不存在「包含」：
    '   VCC1x Like VCC# = False，不需要也不会写 ^ $
    ' ★ * 会吃掉所有东西，所以 "VCC" Like "*VCC*" 才是「包含 VCC」

    ' ---- 大小写：默认区分 ----
    MsgBox "vcc1 像 VCC#：" & ("vcc1" Like "VCC#")         ' False
    ' ★ 想忽略大小写：两边都 UCase$，或模块顶部 Option Compare Text
    '   （后者影响整个模块，慎用；见 DemoSearch）

    ' ---- 特殊字符要放进 [] 里才按字面匹配 ----
    MsgBox "a?b 像 a[?]b：" & ("a?b" Like "a[?]b")          ' True
    MsgBox "a*b 像 a[*]b：" & ("a*b" Like "a[*]b")          ' True
    MsgBox "a#b 像 a[#]b：" & ("a#b" Like "a[#]b")          ' True
    MsgBox "a[b 像 a[[]b：" & ("a[b" Like "a[[]b")          ' True
    ' ★ 右方括号 ] 在组外就是普通字符：a]b Like a]b = True

    ' ---- 字符列表与区间 ----
    MsgBox "VCC 像 [VGC]*：" & ("VCC" Like "[VGC]*")        ' True
    MsgBox "RESET 像 [VGC]*：" & ("RESET" Like "[VGC]*")    ' False
    MsgBox "A 像 [A-Za-z]：" & ("A" Like "[A-Za-z]")        ' True
    MsgBox "5 像 [!0-9]：" & ("5" Like "[!0-9]")            ' False

    ' ---- 实战校验：引脚名（大写字母开头，后面字母/数字/下划线）----
    MsgBox "VCC_3V3 合法：" & ("VCC_3V3" Like "[A-Z][A-Z0-9_]*")    ' True
    MsgBox "3V3 合法：" & ("3V3" Like "[A-Z][A-Z0-9_]*")            ' False

    ' ---- 版本号：# 是「恰好一个数字」，别当通配符 ----
    MsgBox "1.25V 像 #.##V：" & ("1.25V" Like "#.##V")      ' True
    MsgBox "12.25V 像 #.##V：" & ("12.25V" Like "#.##V")    ' False

    ' ---- 空串 ----
    MsgBox "空串像 *：" & ("" Like "*")        ' True（* 允许 0 个字符）
    MsgBox "空串像 ?：" & ("" Like "?")        ' False
    ' [] 表示空串，极少用："" Like "[]" = True

    ' ★ Like 不是正则：
    '   没有 | 或运算、没有 {2,3} 计数、没有分组捕获、没有锚点概念；
    '   [A-Z] 只代表「一个字符」。复杂规则（时间戳、邮箱、带小数的电压）
    '   用下面的 DemoRegex
End Sub


' ============================================================
' 九、数值 ↔ 字符串：CStr / Str / Val / Format
' ============================================================
Sub DemoNumberString()
    MsgBox "CStr：" & CStr(3.3)                        ' 3.3（看区域设置，可能是 3,3）
    MsgBox "Str 留符号位：[" & Str$(3.3) & "]"          ' [ 3.3]，正数前有空格
    MsgBox "Str 负数：[" & Str$(-3.3) & "]"             ' [-3.3]
    ' ★ 用 Str 拼日志会多出前导空格，日志里用 CStr 或 Format$ 更顺眼
    ' ★ 但 CStr / Format 受区域设置影响（3.3 可能变 3,3），
    '   而 Str 的小数点固定是英文句点 —— 拼 CSV / 传给别的程序用 Trim$(Str$(x)) 最稳

    ' ---- Val：从字符串里抠数字，只认英文小数点 ----
    MsgBox "Val(""3.310V"")：" & Val("3.310V")          ' 3.31
    MsgBox "Val(""3,310"")：" & Val("3,310")            ' 3 ← 遇逗号就停
    MsgBox "Val(""$3.3"")：" & Val("$3.3")              ' 0 ← 不认货币符号
    MsgBox "Val(""&HFF"")：" & Val("&HFF")              ' 255 ← 认十六进制
    MsgBox "Val(""abc"")：" & Val("abc")                ' 0
    ' ★ Val 会跳过空格/Tab/换行，遇到第一个不认识的字符就停；
    '   "3,310" 这种千分位必须先 Replace 掉逗号（见 1.6 的 ToDouble）
    ' ★ Val 永远按英文句点解析，跟区域设置无关 —— 这就是解析数据文件
    '   要用 Val 而不是 CDbl 的原因（见 1.7）

    ' ---- IsNumeric：宽松，别拿它当严格校验 ----
    ' ★ 它认可千分位和日期（见 README 踩坑清单第 9 条），
    '   想确认「是不是纯数字」用 Like 或正则更可靠

    ' ---- Format$：补零、千分位、百分比、日期 ----
    MsgBox "补小数：" & Format$(3.3, "0.000")              ' 3.300
    MsgBox "数值补零：" & Format$(12, "0000")              ' 0012
    MsgBox "千分位：" & Format$(1234.5, "#,##0.00")        ' 1,234.50
    MsgBox "百分比：" & Format$(0.256, "0.0%")             ' 25.6%
    MsgBox "日期：" & Format$(Now, "yyyy-mm-dd hh:nn:ss")  ' 2026-10-27 09:30:00
    ' ★ 分钟是 nn 不是 mm（mm 是月份），VBA 里最常见的日期格式错
    ' ★ Format 全部受区域设置影响：小数点、千分位、日期分隔符都可能变，
    '   写进 CSV / 文件名前先想清楚目标机器认不认（见 1.7）

    ' ---- 文本补位：VBA 没有 padStart / padEnd ----
    MsgBox "文本左补：[" & PadLeft("7", 3, "0") & "]"      ' [007]
    MsgBox "文本右补：[" & PadRight("VCC", 5) & "]"        ' [VCC  ]
    ' ★ 数值补位优先用 Format$（不用自己转字符串再补）；文本用 PadLeft/PadRight
    ' ★ 补位宽度按 Len（字符数）算，中文一个字符算 1

    ' ---- 往返：数值 → 字符串 → 数值 ----
    MsgBox "往返：" & Val(Trim$(Str$(3.31)))               ' 3.31
End Sub


' ============================================================
' 十、大量拼接：& 在循环里反复拼接是 O(n²)
'   几千行日志这样写会明显变慢；用数组收集 + Join 一次成型
' ============================================================
Sub DemoConcat()
    ' ---- & 永远拼接，+ 会先尝试加法 ----
    Dim a As Variant, b As Variant
    a = "1": b = "2"
    MsgBox "& 拼接：" & (a & b)          ' 12
    MsgBox "+ 相加：" & (a + b)          ' 3
    ' ★ Variant 里装着数字字符串时 + 走加法；两个都是 String 类型时 + 才是拼接。
    '   别赌这个区别，拼接一律用 &
    ' ★ 数值参与 & 时按 CStr 规则转换，受区域设置影响（3.3 可能变 3,3），
    '   拼进 CSV 前先 Format$ / Trim$(Str$())

    ' ---- Null 与空串 ----
    MsgBox "[" & Null & "]"              ' []，& 把 Null 当空串
    ' MsgBox "[" + Null + "]"            ' 整个表达式变 Null → 报错 94
    ' ★ 从单元格读出来的值可能是 Null（公式返回空）或 Empty（空单元格），
    '   拼之前用 v & "" 兜底

    ' ---- 数组收集 + Join ----
    Dim lines() As String
    Dim i As Long
    ReDim lines(0 To 999)

    For i = 0 To 999
        lines(i) = "2026-10-26 09:30:00,VCC," & CStr(i) & ",Pass"
    Next

    Dim text As String
    text = Join(lines, vbCrLf)           ' 一次性拼好
    MsgBox "总长度：" & Len(text) & " 字符"

    ' ★ 反例：text = text & newLine 写 1000 次，每次都要重新分配整块内存，
    '   数据量上万行时耗时能差出几倍
    ' ★ ReDim Preserve 每轮扩容也是 O(n²)：能预估就先开够，
    '   估不准就按 2 倍扩容（1024 → 2048 → 4096）
End Sub


' ============================================================
' 十一、正则：Like 扛不住时上 VBScript.RegExp
'   后期绑定，不需要在「工具 → 引用」里勾任何东西
' ============================================================
Sub DemoRegex()
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True        ' ★ 默认 False：只处理第一个匹配
    re.IgnoreCase = True    ' ★ 默认 False：区分大小写
    re.MultiLine = False    ' True 时 ^ $ 按「每行」算

    ' ---- 判断：Test 返回 True/False ----
    re.Pattern = "(\d+(?:\.\d+)?)\s*V"
    MsgBox "VCC 3.31V 有电压吗：" & re.Test("VCC 3.31V")    ' True
    MsgBox "GND 0V 有电压吗：" & re.Test("GND 0V")           ' True

    ' ---- 取值：Execute → MatchCollection ----
    Dim mc As Object, m As Object
    Set mc = re.Execute("VCC 3.31V / GND 0V / CLK 1.8V")
    MsgBox "匹配到 " & mc.Count & " 处"                      ' 3（Global 必须为 True）
    For Each m In mc
        Debug.Print m.Value; "  位置(0基)="; m.FirstIndex; _
                    "  数值="; m.SubMatches(0)
    Next
    ' ★ FirstIndex 是 0 基（和 VBA 其它位置函数相反），换算成 VBA 位置要 +1
    ' ★ SubMatches 也是 0 基；没匹配到的分组返回空串
    ' ★ Global = False 时 Execute 只返回第一个匹配，Count 永远是 1

    ' ---- 替换：$1/$2 引用分组，$0 是整个匹配 ----
    re.Pattern = "\s+"
    MsgBox "压空格：[" & re.Replace("  VCC   3.31V  ", " ") & "]"     ' [ VCC 3.31V ]
    re.Pattern = "(\w+),(\w+)"
    MsgBox "换列：" & re.Replace("VCC,Pass", "$2,$1")                 ' Pass,VCC
    ' ★ 只替换第一次：re.Global = False

    ' ---- 校验：^ $ 锚定整串 ----
    re.IgnoreCase = False                             ' 属性随时可改，不用重新 CreateObject
    re.Pattern = "^[A-Z][A-Z0-9_]*$"
    MsgBox "VCC_3V3 合法：" & re.Test("VCC_3V3")     ' True
    MsgBox "3V3 合法：" & re.Test("3V3")             ' False
    MsgBox "vcc_3v3 合法：" & re.Test("vcc_3v3")     ' False（大写校验）

    ' ---- 按正则拆：没有 split，先统一成哨兵字符再 Split ----
    Dim s As String, parts As Variant
    s = "VCC 3.31V;GND,0V"
    re.Pattern = "\s*[;,]\s*"
    parts = Split(re.Replace(s, ","), ",")
    MsgBox "拆成 " & (UBound(parts) + 1) & " 段：" & Join(parts, " | ")

    ' ★ VBScript 正则的能力边界：
    '   ✔ 支持 \d \w \s \b、分组 () (?:)、量词 * + ? {n,m}、懒惰 *?、
    '     前瞻 (?=) (?!)
    '   ✘ 不支持后顾 (?<=) (?<!)、命名分组 (?<name>)、Unicode 类 \p{L}、
    '     内联标志 (?i)（用 IgnoreCase 代替）
    ' ★ \w 只含 ASCII 字母数字下划线，中文不算 \w；匹配「任意非空白」用 [^\s]
    ' ★ . 不匹配换行；跨行整体匹配用 [\s\S]，按行锚定用 MultiLine = True
    ' ★ Pattern 来自用户输入 / 配置文件时先兜错，写错会报 5017（正则语法错误）
    ' ★ 循环里别反复 CreateObject：Set 一次、改 Pattern 复用，
    '   或用 Static re As Object 缓存（见 3.4）
End Sub


' ============================================================
' 十二、ATE 实用解析：从一条日志里取出 时间 / 引脚 / 电压 / 结果
'   固定 4 段就用 Split；段数不固定再考虑 InStr + Mid
' ============================================================
Sub DemoParseLog()
    Dim line As String
    line = "2026-10-26 09:30:00,VCC,3.310,Pass"
    ' 时间本身含空格但没有逗号，所以直接 Split 是安全的
    ' （若时间格式写成 "2026-10-26 09:30:00" 之外的形式，先确认字段里有没有分隔符）

    Dim f As Variant
    f = Split(line, ",")
    If UBound(f) - LBound(f) + 1 < 4 Then
        MsgBox "字段不足，跳过该行：" & line
        Exit Sub
    End If

    Dim t As String, pin As String, volt As Double, res As String
    t = f(0)
    pin = f(1)
    volt = Val(f(2))                          ' Val 只认英文小数点，避免区域设置坑
    res = f(3)
    MsgBox t & " / " & pin & " / " & volt & " V / " & res

    ' ★ 真实数据里更麻烦的三种情况，遇到就得上「带引号的 CSV 解析」（见 5.2）：
    '   1) 字段里含逗号，如 "VCC,3V3"
    '   2) 字段用引号包着，如 "3,310"
    '   3) 结尾有 \r（从 Windows 文本读出来时），会粘在最后一个字段上
End Sub


' ============================================================
' 工具函数：VBA 没有的那几个
'   纯函数，直接调；想限制在本模块用，在前面加 Private
'   以后抽公共模块 Util.bas 时一起搬过去
' ============================================================
Function StartsWith(ByVal s As String, ByVal prefix As String) As Boolean
    If Len(prefix) > Len(s) Then Exit Function           ' 没赋值 → 返回 False
    StartsWith = (StrComp(Left$(s, Len(prefix)), prefix, vbBinaryCompare) = 0)
End Function

Function EndsWith(ByVal s As String, ByVal suffix As String) As Boolean
    If Len(suffix) > Len(s) Then Exit Function
    EndsWith = (StrComp(Right$(s, Len(suffix)), suffix, vbBinaryCompare) = 0)
End Function

' 左补位：PadLeft("7", 3, "0") = "007"
' ★ 超过宽度原样返回，不会截断；pad 传多字符只取第一个
Function PadLeft(ByVal s As String, ByVal width As Long, _
                 Optional ByVal pad As String = " ") As String
    If Len(pad) = 0 Then pad = " "                        ' String$ 不认空串
    If Len(s) >= width Then
        PadLeft = s
    Else
        PadLeft = String$(width - Len(s), pad) & s
    End If
End Function

Function PadRight(ByVal s As String, ByVal width As Long, _
                  Optional ByVal pad As String = " ") As String
    If Len(pad) = 0 Then pad = " "
    If Len(s) >= width Then
        PadRight = s
    Else
        PadRight = s & String$(width - Len(s), pad)
    End If
End Function

' 去首尾的任意指定字符：TrimChars("""VCC""", """") = VCC
' 和 Trim 一样只看首尾，中间不动；chars 里可以一次列多个字符
Function TrimChars(ByVal s As String, ByVal chars As String) As String
    Dim i As Long, j As Long
    i = 1
    Do While i <= Len(s)
        If InStr(chars, Mid$(s, i, 1)) = 0 Then Exit Do
        i = i + 1
    Loop
    j = Len(s)
    Do While j >= i
        If InStr(chars, Mid$(s, j, 1)) = 0 Then Exit Do
        j = j - 1
    Loop
    TrimChars = Mid$(s, i, j - i + 1)
End Function

' 按多个分隔符拆：SplitAny("VCC,GND;CLK", ",;") = 3 段
' ★ 连续分隔符照样会拆出空串（和 Split 行为一致，不自动合并）
Function SplitAny(ByVal s As String, ByVal delims As String) As Variant
    Dim sentinel As String, i As Long
    sentinel = Chr(31)                                    ' 单元分隔符，正常文本里没有
    For i = 1 To Len(delims)
        s = Replace(s, Mid$(delims, i, 1), sentinel)
    Next
    SplitAny = Split(s, sentinel)
End Function

' 判空：只由空白组成的串也算空
Function IsBlank(ByVal s As String) As Boolean
    IsBlank = (Len(Trim$(s)) = 0)
End Function
