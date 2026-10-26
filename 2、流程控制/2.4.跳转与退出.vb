Option Explicit

' ============================================================
' 2.4 跳转与退出：提前离开，减少嵌套
'     Exit Sub / Exit Function   离开整个过程
'     Exit For / Exit Do         离开当前这一层循环（2.2 / 2.3 已见）
'     GoTo 标签                  跳到同一过程内的另一个位置
'     GoTo 只能在本过程内跳，不能跳进循环内部、不能跳到别的模块
'
' ★ 拿 JS 建立坐标：VBA 根本没有 return / break / continue 这三个关键字，
'   全部由 Exit 系列 + GoTo 顶替，先记住这张对照表：
'
'     JS                                  VBA
'     --------------------------------    ----------------------------------
'     function f() { ... return; }        Exit Sub
'     function f() { ... return x; }      f = x 然后 Exit Function（没有 return 语句）
'     break;                              Exit For / Exit Do
'     continue;                           没有对应物 → GoTo 标签，或循环体首部 If 挡掉
'     throw new Error(...)                Err.Raise（见 3.3）
'     try { ... } finally { ... }         On Error GoTo Cleanup ... Cleanup:（本文件下半段）
'     goto（JS 里压根没有）               GoTo，只留给"跳到过程末尾做善后"
'
' ★ 两个共同的限制：
'   1) Exit 无论哪种，都只能退一层；也没有标签式 break outer
'   2) 想离开整个过程就用 Exit Sub / Exit Function，它比逐层 Exit 干净
' ============================================================
Sub DemoJump()

    ' ---- 卫语句（guard clause）：不合格就提前 Exit Sub ----
    ' 两个样例对比看效果：第一个走完全程，第二个在卫语句处就结束
    CheckPin 3.3
    CheckPin 4.2

    ' ---- GoTo 的正当用法：统一收尾 ----
    CleanupDemo

    ' ---- 结论 ----
    ' 分支用 If / Select Case，循环提前结束用 Exit For / Exit Do，
    ' 过程提前结束用 Exit Sub / Exit Function；
    ' GoTo 只留给"跳到末尾做善后"这一种场景
    '
    ' 翻成 JS 就是三行：return; / break; / try { } finally { }
    ' 也就是说日常九成的"提前离开"在 VBA 里都有现成的对应物，
    ' 真正非 GoTo 不可的只剩 try/finally 那一种，其余的别用 GoTo
End Sub


' ============================================================
' 卫语句示范：先挡掉不合格的情况，主逻辑不用层层缩进
' 不合法 → 直接 Exit Sub，后面的语句根本不执行
'   JS 写法一模一样：if (!judge(v)) { log(...); return; }
'   VBA 只把 return 换成 Exit Sub 这一处
' ============================================================
Sub CheckPin(ByVal v As Double)
    If Not Judge(v, 3#, 3.6) Then
        MsgBox v & " V 超限 → Fail，本次测试到此结束"
        Exit Sub                              ' ← 等价 JS 的 return;（Sub 没有返回值）
    End If
    MsgBox v & " V 合格 → 继续下一项测试"
End Sub
' ★ 好处和 JS 一样：不合格的情况在顶部一次性挡掉，主逻辑不用层层缩进，
'   也不会因为后面新增分支而漏掉某条路径


Sub CleanupDemo()
    ' 中间任何一步出错都跳到 Cleanup 做善后，
    ' 清理代码只写一份，不会因为新增分支而漏掉（第 4 章的提速开关就靠它恢复）
    '
    ' JS 的等价物是 try/catch/finally，只是 VBA 把它压成了一个标签：
    '   try { ...干活... } catch (e) { ... } finally { ...善后... }
    '   ↓
    '   On Error GoTo Cleanup : ...干活... : Cleanup: ...善后...
    ' 差别：VBA 的 Cleanup 是"正常路径也会掉进去"的落点（像 finally 一定会执行），
    '       JS 的 catch 只接错误；所以这里正常跑完也要保证善后代码能安全重入
    Dim state As String
    state = "未开始"
    On Error GoTo Cleanup                     ' 3.3 会细讲 On Error

    state = "处理中"
    ' ... 真正干活的代码 ...

Cleanup:
    On Error Resume Next
    ' 这里做善后：恢复 ScreenUpdating、关闭文件、写日志 ...
    MsgBox "状态：" & state & " → 已收尾"
    ' ★ 标签后面带冒号、顶格写；标签名在同一个过程内不能重名
    ' ★ 标签只是"位置标记"：程序正常往下走时也会穿过它，
    '   所以正常路径要么 Exit Sub 提前离开，要么像上面这样让收尾能安全重入
    ' ★ JS 读者可以类比 switch 的 fall-through：标签不拦流程，落到哪就从哪继续往下走；
    '   JS 没有 goto，这段逻辑在 JS 里只能写成 try/finally，或者抽一个 cleanup() 函数
End Sub


' ============================================================
' 判定函数：返回值是 Boolean，默认值 False，
' 所以"提前返回"时什么都不用写，直接 Exit Function
'
' JS 里是这样：
'   function judge(v, lo, hi) {
'       if (v < lo) return false;
'       if (v > hi) return false;
'       return true;
'   }
'
' VBA 的差别只有一处，但很关键：没有 return 语句，返回值靠"给函数名赋值"。
'   函数名  = 值        相当于 return 值
'   Exit Function       相当于 return（此时返回函数名的当前值）
' 函数没赋值就退出时，返回的是该类型的默认值：Boolean→False、数值→0、String→""
' 所以上面两个提前返回连 false 都不用写 —— 但也意味着"忘了给函数名赋值"
' 不会报错，只会静默返回 False（这是 VBA 里最常见的隐形 bug，见 README 第 14 条）
' ============================================================
Function Judge(ByVal v As Double, ByVal lo As Double, ByVal hi As Double) As Boolean
    If v < lo Then Exit Function              ' ← return false
    If v > hi Then Exit Function              ' ← return false
    Judge = True                              ' ← return true（先赋值）
End Function                                   '   隐式返回，不需要再写 Exit Function


' ============================================================
' 附：本文件走一遍 —— 从调用到返回的完整执行流程
'   阅读顺序（从上往下）≠ 执行顺序（从入口进去、跳出来、再跳进去）
'   建议对照着按 F8 单步走一遍，比看注释快
' ============================================================
'
' 【第 0 步】从哪儿开始
'   光标放进 DemoJump 里按 F5，只执行这一个 Sub；
'   其它 Sub / Function 不会自动跑，它们只在被调用时才进来。
'   Sub 跑完就结束、没有返回值，Function 则要把值交回去（这就是两者的分界）。
'
' 【第 1 步】DemoJump 第 1 行：CheckPin 3.3
'   ① 传值：3.3 是 Double 字面量
'   ② 接受参数：CheckPin(ByVal v As Double) → ByVal 把 3.3 复制一份给 v，
'      子过程里怎么改 v 都影响不到外面
'   ③ 调用写法：不带 Call 就不能加括号（多参数时加括号会报"应为 ="，
'      细节见 3.1）；写 Call CheckPin(3.3) 也完全正确
'   ④ 进入 CheckPin 体内，从第一行往下执行
'
' 【第 2 步】CheckPin 里调 Judge(v, 3#, 3.6)
'   ① 实参 v = 3.3；3# 是 Double 字面量（# 是类型声明字符，见 1.2）
'   ② Judge 三个形参都是 ByVal，各自复制一份，互不干扰
'   ③ Judge 体内逐句判断：
'        v < lo ?   3.3 < 3     否
'        v > hi ?   3.3 > 3.6   否
'        Judge = True           ← 给函数名赋值 = 设定返回值
'        走到 End Function      → 把 True 交回调用处
'   ④ 回到 CheckPin：If Not True Then → False，条件不成立，跳过 Exit Sub
'   ⑤ 打印"合格" → End Sub → 返回 DemoJump
'
' 【第 3 步】DemoJump 第 2 行：CheckPin 4.2
'   ① 同样 ByVal，v = 4.2
'   ② Judge(4.2, 3#, 3.6)：v > hi 成立 → 执行 Exit Function
'      ★ 这一刻 Judge 还是默认值 False（没走到 Judge = True），
'        所以 Exit Function 直接返回 False —— 这就是"提前返回"，
'        既不用也不能写 return False
'   ③ 回到 CheckPin：Not False = True → 进入 If
'   ④ MsgBox 报超限 → Exit Sub
'      ★ 立刻结束整个 CheckPin，后面那句"合格"永远不会执行
'   ⑤ 返回 DemoJump
'
' 【第 4 步】DemoJump 第 3 行：CleanupDemo
'   ① On Error GoTo Cleanup 先"登记"出事跳哪儿，还没真跳
'   ② 正常路径一路走到底，自然落到 Cleanup: 标签
'      ★ 标签不拦流程，正常路径也会穿过去（效果等同 JS 的 finally）
'   ③ 善后执行完 → End Sub
'
' 【第 5 步】回到 DemoJump，走到 End Sub，整段代码结束
'
' ------------------------------------------------------------
' 同一个错误，三种提示方式（这是"错误提示"的关键）
'   1) 没有任何 On Error
'        → 错误沿调用栈往上冒；一路没人接，Excel 弹出"运行时错误 xxx"，
'          点"调试"停在出错那一行，点"结束"终止全部代码
'   2) On Error Resume Next
'        → 不提示，跳过出错那句继续跑。最危险：表面"跑完了"，
'          实际数据可能缺一大截（所以要立刻 On Error GoTo 0 关掉）
'   3) On Error GoTo 标签
'        → 跳到标签，自己弹 MsgBox、写日志、恢复设置，再决定 Resume 还是退出
'   本文件里：CheckPin / Judge 都没写处理器，所以出错会冒到 DemoJump；
'   CleanupDemo 有处理器，中途出错不弹框，而是走 Cleanup 收尾。
'
' ------------------------------------------------------------
' 一句话记住这条链
'   DemoJump 调 CheckPin，CheckPin 调 Judge；
'   Exit 只结束"自己这一个"过程，然后回到上一层继续；
'   只有没人接的错误才会一路冒到最外面弹框。
' ============================================================
