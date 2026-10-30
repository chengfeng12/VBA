Option Explicit

' ============================================================
' 5.0 入口：本章总入口
'     下面几行调用默认注释掉了，去掉注释即可依次运行
'     也可以把光标放进任意一个 Sub 里按 F5 单独运行
'     ★ 本章只读写临时目录（Environ("TEMP")），不会碰你的数据
' ============================================================
Sub HelloPractical()
    MsgBox "第 5 章：实战基础 —— 文本 / 字典 / COM"

    ' 本章各小节（每个 Sub 在独立模块里）
    ' DemoTrim            ' 5.1 去空白：Trim 只认普通空格 / 全角空格 / Tab / BOM
    ' DemoSearch          ' 5.1 查找比较：InStr / InStrRev / StrComp
    ' DemoSlice           ' 5.1 截取：Left / Mid / Right / Mid 语句
    ' DemoSplitJoin       ' 5.1 拆分合并：Split / Join / Filter
    ' DemoReplace         ' 5.1 替换：起点 / 次数 / 忽略大小写
    ' DemoCase            ' 5.1 大小写与全半角：UCase / LCase / StrConv
    ' DemoCharCode        ' 5.1 字符码：Asc / Chr / StrReverse / String$
    ' DemoLike            ' 5.1 模式匹配 Like
    ' DemoNumberString    ' 5.1 数值 ↔ 字符串：CStr / Str / Val / Format
    ' DemoConcat          ' 5.1 大量拼接用数组 + Join
    ' DemoRegex           ' 5.1 正则：VBScript.RegExp
    ' DemoParseLog        ' 5.1 解析一条测试日志
    ' DemoFreeFile        ' 5.2 文件号的取用与释放
    ' DemoFileWrite       ' 5.2 写 CSV（造示例文件）
    ' DemoFileRead        ' 5.2 逐行读文件
    ' DemoParseCsv        ' 5.2 带引号的 CSV 解析
    ' DemoListFiles       ' 5.2 遍历目录
    ' DemoFileManage      ' 5.2 文件管理：Kill / FileCopy / Name / MkDir
    ' DemoEnviron         ' 5.2 读环境变量（TEMP / USERNAME 等）
    ' DemoDictBinding     ' 5.3 早期绑定 vs 后期绑定（Dim d As Dictionary）
    ' DemoDictBasics      ' 5.3 字典增删改查
    ' DemoDictChangeKey   ' 5.3 修改键：d.Key(旧)=新、批量规范化
    ' DemoPinLimit        ' 5.3 引脚 → 限值映射
    ' DemoCountByBin      ' 5.3 按 bin 统计
    ' DemoDictTraps       ' 5.3 字典的坑
    ' DemoComBasics       ' 5.4 CreateObject 基础
    ' DemoComTraps        ' 5.4 COM 实战坑（说明为主）
    ' DemoRndBasics       ' 5.5 Rnd / Randomize 基础
    ' DemoMockData        ' 5.5 造 100 颗模拟数据算良率
    ' DemoFixedSeed       ' 5.5 固定种子复现
    ' DemoRndTraps        ' 5.5 Rnd 的反直觉行为
End Sub
