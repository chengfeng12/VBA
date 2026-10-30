Option Explicit

' ============================================================
' 第 5 章  实战基础（文本 / 字典 / COM）
' ------------------------------------------------------------
' 为什么单独成章：
'   前 4 章是"语言本身"，这一章是"每天真正要干活的工具"。
'   量产脚本 90% 的时间花在四件事上 ——
'     读文件、切字符串、查表统计、跟别的程序打交道
'   这四件事在前 4 章里要么完全没有，要么只有一两行注释。
' ------------------------------------------------------------
' 用法：每个 .vb 文件 = 一个独立的 VBA 标准模块，粘进工程即可
'       本章示例只动临时目录（Environ("TEMP")），不会碰你的数据
' ------------------------------------------------------------
' 目录
'   5.0.入口.vb               HelloPractical —— 本章总入口
'   5.1.字符串处理.vb         Trim / InStr / InStrRev / StrComp / Left / Mid /
'                             Right / Split / Join / Filter / Replace / Like /
'                             Asc / Chr / StrConv / Format / 正则 / 拼接性能
'                             + StartsWith / EndsWith / PadLeft / TrimChars 等工具函数
'   5.2.文本文件与CSV.vb      Open / Line Input / Print # / FreeFile / EOF /
'                             Dir 遍历、带引号 CSV 解析、UTF-8 读写
'   5.3.Dictionary字典.vb     早期/后期绑定、CreateObject、Add/Exists/Keys、
'                             改键（Key 属性）、计数与映射
'   5.4.COM与CreateObject.vb  后期绑定、常用 ProgID、Excel.Application、
'                             释放与防泄漏
'   5.5.随机与模拟数据.vb     Rnd / Randomize、整数区间、标称±容差造数据、
'                             固定种子复现
' ------------------------------------------------------------
' 和 JS 的整体对照（迁移时按这张表找）
'
'   读文件           fs.readFileSync        Open For Input + Line Input
'   写文件           fs.writeFileSync       Open For Output + Print #
'   切字符串         s.split / s.slice      Split / Mid$ / Left$ / Right$
'   查表             new Map()              CreateObject("Scripting.Dictionary")
'   调外部程序       child_process          CreateObject("WScript.Shell")
'   判断浮点相等      Math.abs(a-b) < eps    Abs(a-b) < eps（见 1.7）
'   释放资源         GC 自动                Set obj = Nothing（进程外对象必须手动）
' ------------------------------------------------------------
' 先记住这几条
'   1. 量产优先用 CSV + 文本流，别用 Excel COM 打开数据文件
'   2. VBA 原生 Open 读写的是 ANSI；UTF-8 中文要用 ADODB.Stream
'   3. Split 得到的是 0 基数组；InStr 找不到返回 0（不是 -1）
'   4. Dictionary 取值前先 Exists —— 直接 d("新键") 会顺手建一个空条目
'   5. 进程外对象（Excel.Application）必须 Quit + Set Nothing，否则残留进程
'   6. 字段里可能含逗号，Split 只够应付干净数据；带引号要用状态机解析
'   7. Rnd 不调 Randomize 就"永远是同一条序列"（和 JS 完全不同）；
'      要复现数据反而故意用固定种子 Rnd(-1)
' ============================================================

' 本模块只放说明，没有可执行的 Sub
