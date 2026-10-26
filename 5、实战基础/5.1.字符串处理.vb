Option Explicit

' ============================================================
' 5.1 字符串处理：引脚名、版本号、日志、结果拼接天天要用
' ------------------------------------------------------------
' 和 JS 对照（几乎一一对应，只有下标规则不同）：
'     s.trim()              → Trim$(s)
'     s.indexOf("x")        → InStr(s, "x")      找到返回 1 基位置，找不到返回 0
'     s.slice(1, 4)         → Mid$(s, 2, 3)      第二个参数是"起点"，1 基
'     s.slice(0, 3)         → Left$(s, 3)
'     s.slice(-3)           → Right$(s, 3)
'     s.split(",")          → Split(s, ",")      得到 0 基数组
'     arr.join(",")         → Join(arr, ",")
'     s.replaceAll(a, b)    → Replace(s, a, b)
'     s.toUpperCase()       → UCase$(s)
'     /^VCC\d+$/            → s Like "VCC##"
'     s1 === s2             → StrComp(s1, s2, vbBinaryCompare) = 0
' ============================================================
Sub DemoString()
    Dim s As String
    s = "  VCC_3V3_1.25V  "

    ' ---- 去空格 ----
    MsgBox "[" & Trim$(s) & "]"                      ' [VCC_3V3_1.25V]
    ' ★ Trim 只处理普通空格（Chr(32)），全角空格和 Tab 一律不管：
    '   从别的系统导出的文件里常混全角空格（Chr(12288)），看着一模一样，
    '   但 InStr / 比较全都不认 → 先 Replace 再 Trim：
    '   s = Trim$(Replace(s, Chr(12288), " "))

    ' ---- 查找：找不到返回 0，不是 -1（JS 是 -1）----
    MsgBox "InStr 找 V3：" & InStr(s, "V3")           ' 8
    If InStr(s, "GND") = 0 Then MsgBox "没找到 GND"
    ' ★ 从右往左找用 InStrRev；带起点：InStr(3, s, "V") 表示从第 3 个字符开始

    ' ---- 截取：越界不报错，返回空串 ----
    MsgBox "Left 3：" & Left$(s, 3)                   ' "  V"
    MsgBox "Mid 从 3 取 5：" & Mid$(s, 3, 5)          ' "VCC_3"
    MsgBox "Right 4：" & Right$(Trim$(s), 4)          ' "25V"
    MsgBox "Mid 越界：" & "[" & Mid$("abc", 10, 2) & "]"   ' []，不报错
    ' ★ 越界静默返回空串，比 JS 的 undefined 更隐蔽，取完记得判 Len

    ' ---- 拆分与合并：Split 的数组是 0 基 ----
    Dim parts As Variant, i As Long
    parts = Split("VCC,3.31,Pass", ",")
    MsgBox "段数：" & (UBound(parts) + 1)              ' 3
    For i = LBound(parts) To UBound(parts)
        Debug.Print "第 " & i + 1 & " 段：" & parts(i)  ' 0 基 → 打印时 +1
    Next
    MsgBox "合回去：" & Join(parts, " | ")             ' VCC | 3.31 | Pass
    ' ★ Split("", ",") 返回"含 1 个空串"的数组（UBound = 0，不是 -1），
    '   先判 Len(s) = 0 再去取 parts(0)，否则会拿到空串当有效值

    ' ---- 替换：默认全替换、区分大小写 ----
    MsgBox Replace("VCC-vcc", "vcc", "GND")           ' "VCC-GND"
    MsgBox Replace("VCC-vcc", "vcc", "GND", 1, -1, vbTextCompare)   ' "GND-GND"
    ' 完整参数：Replace(原串, 找什么, 换成什么, [起始位置], [次数], [比较模式])
    '   次数 -1 = 全部替换；比较模式 vbBinaryCompare(0) 区分大小写 / vbTextCompare(1) 不区分
    ' ★ 这两个常量在 StrComp、InStr、字典的 CompareMode 里都在用，值就是 0 和 1，
    '   但写名字比写数字清楚得多

    ' ---- 大小写 ----
    MsgBox UCase$("pass") & " / " & LCase$("FAIL")     ' PASS / fail
    ' ★ 比较时别先 UCase 再比，直接用 StrComp(..., vbTextCompare) 更快更清楚

    ' ---- 模式匹配：Like（正则的轻量替代）----
    '   ?  任意单字符      *  任意多个字符      #  任意单个数字
    '   [A-Z] 字符区间     [!0-9] 取反
    MsgBox "VCC12 匹配 VCC##：" & ("VCC12" Like "VCC##")        ' True
    MsgBox "VCC1A 匹配 VCC##：" & ("VCC1A" Like "VCC##")        ' False
    MsgBox "版本号匹配合格：" & ("1.25V" Like "#.##V")           ' True
    ' ★ Like 的 * 会吃掉所有东西，写 "^VCC" 这种正则习惯在这里行不通
    '   （VBA 的 Like 没有锚点概念），要精确匹配就得把整串写全

    ' ---- 数值 ↔ 字符串 ----
    MsgBox "CStr：" & CStr(3.3)                        ' "3.3"
    MsgBox "Str 会留符号位：[" & Str$(3.3) & "]"        ' [ 3.3] ← 正数前面有空格
    ' ★ 用 Str 拼日志会多出前导空格，一律用 CStr 或 Format$
    MsgBox "补零：" & Format$(3.3, "0.000")             ' 3.300
    ' ★ Format 受系统区域设置影响：小数点可能是逗号的系统上会输出 "3,300"，
    '   拼进 CSV 就多一列 —— 见 1.7
    MsgBox "重复字符：" & String$(3, "0")               ' 000
    MsgBox "补空格对齐：[" & Space(5) & "]"              ' [     ]

    ' ---- 长度：Len 是字符数，LenB 是字节数 ----
    MsgBox "Len：" & Len("VCC电压") & "，LenB：" & LenB("VCC电压")   ' 5 / 8
    ' ★ 中文在 GBK 代码页下 Len = 1、LenB = 2；
    '   校验"字段是否超长"时先想清楚是按字符算还是按字节算，
    '   长度和编码绑在一起，换代码页会变
End Sub


' ============================================================
' 逐行拼接的大坑：& 在循环里反复拼接是 O(n²)
'   几千行日志这样写会明显变慢；用数组收集 + Join 一次成型
' ============================================================
Sub DemoConcat()
    Dim lines() As String
    Dim i As Long
    ReDim lines(0 To 999)

    For i = 0 To 999
        lines(i) = "2026-10-26 09:30:00,VCC," & CStr(i) & ",Pass"
    Next

    Dim text As String
    text = Join(lines, vbCrLf)                ' 一次性拼好
    MsgBox "总长度：" & Len(text) & " 字符"
    ' ★ 反例：text = text & newLine 写 1000 次，每次都要重新分配整块内存，
    '   数据量上万行时耗时能差出几倍
End Sub


' ============================================================
' ATE 实用解析：从一条日志里取出 时间 / 引脚 / 电压 / 结果
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

    ' ★ 真实数据里更麻烦的三种情况，遇到就得上"带引号的 CSV 解析"（见 5.2）：
    '   1) 字段里含逗号，如 "VCC,3V3"
    '   2) 字段用引号包着，如 "3,310"
    '   3) 结尾有 \r（从 Windows 文本读出来时），会粘在最后一个字段上
End Sub
