Attribute VB_Name = "C公式校验"
Private SforbidCol As Integer    '禁用标记列
Private CforbidCol As Integer    '禁用标记列
Private Const COMPLEX_TARGET_INDEX As Integer = 3  '复杂校验指标代码的索引位置
Private Const COMPLEX_FRE_INDEX As Integer = 5
Private Const COMPLEXT_SPLIT_COUNT As Integer = 4  '复杂校验指标的分隔符数量
Private Const dbl_Epsilon As Double = 2.22044604925031E-13
Private arrComplexeKeys  As Variant

Public Sub CheckData_ComplexRef()
    If glo_org.count = 0 Then
        Exit Sub
    End If
    If dComplexRef.count = 0 Then
        Exit Sub
    End If
    
    Dim orgs As Variant, i As Integer, org_code As String
    Dim i_rule As Integer
    Dim rule As String, rule_val As String, is_only_err As Boolean
    is_only_err = False
    If gloCft.GetConfig("是否仅输出校验错误项目") = "是" Then is_only_err = True
    
    
    orgs = glo_org.keys
    Dim v_rules As Variant, d_temp As Object, drel_temp As Object, d_rel As Object, v As Integer
    v_rules = dComplexRef.items
    For i = 0 To UBound(orgs)
        ' 遍历所有规则
        org_code = CStr(orgs(i))
        
        For v = 0 To UBound(v_rules)
            Set d_temp = v_rules(v)
            If CStr(d_temp("校验规则")) <> "" Then
                
                Set drel_temp = CreateObject("scripting.dictionary")
                drel_temp("机构代码") = org_code
                drel_temp("校验表单") = CStr(d_temp("校验表单"))
                drel_temp("规则类型") = CStr(d_temp("规则类型"))
                drel_temp("触发指标") = CStr(d_temp("触发指标"))
                '计算结果
                rule = ChangeRule(CStr(d_temp("校验规则")), org_code)
                Set d_rel = yc_Evaluate(rule, ycV(d_temp("thd值")), True)
                If InStr(rule, "thd") > 0 Then rule = replace(rule, "thd", ycV(d_temp("thd值")))
                
                drel_temp("是否执行异常") = d_rel("是否执行异常")
                drel_temp("异常说明") = d_rel("异常说明")
                drel_temp("校验结果") = d_rel("逻辑结果")
                
                If VarType((drel_temp("校验结果"))) = vbBoolean Then
                    If CBool(d_temp("取反标识")) = True Then
                        If CBool(drel_temp("校验结果")) = False Then
                            drel_temp("校验结果") = True
                        Else
                            drel_temp("校验结果") = False
                        End If
                    End If
                    If drel_temp("校验结果") = True Then
                        drel_temp("校验结果") = "错误"
                    Else
                        drel_temp("校验结果") = "正确"
                    End If
                Else
                    drel_temp("校验结果") = "错误"
                End If
                If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
                
                    drel_temp("校验规则") = CStr(d_temp("校验规则"))
                    drel_temp("校验描述") = CStr(d_temp("校验描述"))
                    drel_temp("校验过程") = rule
                    
                    drel_temp("计算值") = ""
                    drel_temp("计算值过程") = ""
                    If CStr(d_temp("计算值公式")) <> "" Then
                        rule_val = ChangeRule(CStr(d_temp("计算值公式")), org_code)
                        Set d_rel = yc_Evaluate(rule_val, False)
                        drel_temp("计算值") = d_rel("逻辑结果")
                        drel_temp("计算值过程") = rule_val
                    End If
                    
                    If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                        Set glo_rel_rule(glo_rel_rule_num) = drel_temp
                    End If
                End If
                glo_rel_rule_num = glo_rel_rule_num + 1
                Set drel_temp = Nothing
            End If
            If v Mod 200 = 0 Then Call gloCft.ShowStatusMessage_Up("正在进行[" & org_code & "]公式校验，执行到第" & CStr(v) & "个。")
        Next
    Next
End Sub



Private Function yc_Evaluate(tRule As String, thd_val As Double, Optional is_lj_val As Boolean = True) As Object
    On Error Resume Next
    
    Dim roundRule As String, d_rel As Object
    roundRule = DoCompare_AddMyComplexExplain_RoundRule(tRule)
    If InStr(roundRule, thd) > 0 Then roundRule = replace(roundRule, "thd", ycV(thd_val))
    If InStr(tRule, thd) > 0 Then roundRule = replace(tRule, "thd", ycV(thd_val))
    
    Set d_rel = CreateObject("scripting.dictionary")
    d_rel("是否执行异常") = "否"
    d_rel("异常说明") = ""
    d_rel("逻辑结果") = ""
    
    If Len(roundRule) < 250 Then
        result = Evaluate(roundRule)
    Else
        Set w_y = gloEt.FindWs(ThisWorkbook, "yc辅助区", 0, True)
        With w_y.Range("A1")
            .ClearContents
            .NumberFormat = "General"
            .Formula = "=" & roundRule
            .Calculate
            result = .value
        End With
    End If
    
     '兜底回退
    If IsError(result) Or err.Number <> 0 Then
        'result = Evaluate(tRule)
        If Len(roundRule) < 250 Then
            result = Evaluate(tRule)
        Else
            Set w_y = gloEt.FindWs(ThisWorkbook, "yc辅助区", 0, True)
            With w_y.Range("A1")
                .ClearContents
                .NumberFormat = "General"
                .Formula = "=" & tRule
                .Calculate
                result = .value
            End With
        End If
    End If
    
    If err.Number <> 0 Then
        d_rel("是否执行异常") = "是"
        d_rel("异常说明") = err.Description
        err.Clear
    Else
        If IsError(result) Then
            d_rel("是否执行异常") = "是"
            d_rel("异常说明") = err.Description
            err.Clear
        ElseIf VarType(result) = vbBoolean Then
            d_rel("逻辑结果") = result
        Else
            d_rel("逻辑结果") = result
            If is_lj_val = True Then
                d_rel("是否执行异常") = "是"
                d_rel("异常说明") = "校验结果并不是逻辑值"
            End If
        End If
    End If
    Set yc_Evaluate = d_rel
End Function
    

'
'   用数据替换校验,
'
Private Function ChangeRule(tRule As String, org_code As String) As String
    Dim reg As Object, mc As Object, m As Object, data_type As String, t_val As Variant
    Dim t_target_code As String, rule As String
    Set reg = CreateObject("VBSCRIPT.REGEXP")
    rule = tRule
    ' 上期
    With reg
        .Global = True: .ignorecase = False: .Pattern = "\{.*?\}"
    End With
    Set mc = reg.Execute(tRule)
    For Each m In mc
        t_target_code = Mid(CStr(m.value), 2, Len(CStr(m.value)) - 2)
        data_type = CStr(Find_Cz(glo_cz_target, t_target_code, "数值类型"))
        t_val = ""
        If glo_d_data.exists(org_code & t_target_code) Then
            t_val = glo_d_data(org_code & t_target_code)("当期值")
        End If
        
        If InStr(data_type, "文") > 0 Then
            rule = replace(rule, CStr(m.value), CStr(t_val))
        Else
            If data_type <> "" Then
                If InStr(data_type, "百") > 0 Or InStr(data_type, "个") > 0 Or InStr(data_type, "户") > 0 Or InStr(data_type, "只") > 0 Then
                    rule = replace(rule, CStr(m.value), ycV(t_val))
                Else
                    rule = replace(rule, CStr(m.value), ycV(t_val))  ' * 10000
                End If
            Else
                rule = replace(rule, CStr(m.value), ycV(t_val))  ' * 10000
            End If
        End If
    Next
    
    '本期
    With reg
        .Global = True: .ignorecase = False: .Pattern = "\[.*?\]"
    End With
    Set mc = reg.Execute(tRule)
    For Each m In mc
        t_target_code = Mid(CStr(m.value), 2, Len(CStr(m.value)) - 2)
        data_type = CStr(Find_Cz(glo_cz_target, t_target_code, "数值类型"))
        t_val = ""
        If glo_d_data.exists(org_code & t_target_code) Then
            t_val = glo_d_data(org_code & t_target_code)("当期值")
        End If
        
        If InStr(data_type, "文") > 0 Then
            rule = replace(rule, CStr(m.value), CStr(t_val))
        Else
            If data_type <> "" Then
                If InStr(data_type, "百") > 0 Or InStr(data_type, "个") > 0 Or InStr(data_type, "户") > 0 Or InStr(data_type, "只") > 0 Then
                    rule = replace(rule, CStr(m.value), ycV(t_val))
                Else
                    rule = replace(rule, CStr(m.value), ycV(t_val))  ' * 10000
                End If
            Else
                rule = replace(rule, CStr(m.value), ycV(t_val))  ' * 10000
            End If
        End If
    Next
    ChangeRule = rule
End Function

Private Sub testround()
    Dim rule As String
    
    arrComplexeKeys = Split(gloCft.GetConfig("系统导数比较-复杂校验-公式精度处理关键字数组"), ",")
    'and(abs(round(28449922100-15634092600-0-474520000-750127600-9831475200-1759706600,2))>round(.01,2),round(28449922100-15634092600-0-474520000-750127600-9831475200-1759706600,2)<0
    'and(or(3.4-3.2>9,4-3.000>4-3),(3+(3.4-3.2))/2-2>4)
    'abs(11097336000-327600-0-11097008400-0)<=.0
    'rule = "and(abs(28449922100-15634092600-0-474520000-750127600-9831475200-1759706600)>.01,28449922100-15634092600-0-474520000-750127600-9831475200-1759706600<0)"
    'rule = "and(abs(28449922100-15634092600-0-474520000-750127600-9831475200-1759706600)>.01,28449922100-15634092600-0-474520000-750127600-9831475200-1759706600>0)"
    '(28449922100-15634092600-0)<(474520000+750127600+9831475200+1759706600)
    'rule = "(28449922100-15634092600-0)<(474520000+750127600+9831475200+1759706600)"
    rule = "andAbS(0.123)<1"
    rule = DoCompare_AddMyComplexExplain_RoundRule(rule)
    Debug.Print (rule)
End Sub

'
'   将两边数值加个round(x,2)
'
Private Function DoCompare_AddMyComplexExplain_RoundRule(rule As String) As String
    On Error GoTo errhandle
    Dim strbk As String, rel As String
    strbk = rule
    
    '去掉空格就方便了
    rule = LCase(replace(rule, " ", ""))
    
    Dim i As Long, strKey As String
    '1 - 存放关键字，2 - 存放 "("
    Dim arr1 As Variant, n1 As Long, arr2 As Variant, n2 As Long, s As Long
    Dim lkh As Integer  '0-非关键字括号,1 - 关键字括号,
    Dim relust As String    '存放位置结果
    Dim zd As Boolean

    relust = ""
    i = 1
    n1 = 0
    n2 = 0
    s = 0
    lkh = 0
    ReDim arr1(1 To (Len(rule) - Len(replace(rule, "(", "")) + 1)) As Variant
    ReDim arr2(1 To (Len(replace(rule, "(", "")) + 1)) As Variant
    
    Do While i <= Len(rule)
        lkh = -1
        zd = False
        strKey = ""
        '遍历rule
        strKey = DoCompare_AddMyComplexExplain_RoundRule_GetKeyType(rule, i)
        If strKey <> "" Then
            '如果是关键字，则读取，并快进
            n1 = n1 + 1
            arr1(n1) = i
            i = i + Len(strKey) - 1
        ElseIf Mid(rule, i, 1) = "(" Then
            '如果是关键字，则读取，并快进
            n2 = n2 + 1
            arr2(n2) = i
        ElseIf DoCompare_AddMyComplexExplain_RoundRule_IsNumber(Mid(rule, i, 1)) Then
            s = s + 1
        Else
            If Mid(rule, i, 1) = ")" Then
                If n1 = 0 And n2 = 0 Then
                    GoTo errhandle
                ElseIf n1 = 0 And n2 > 0 Then
                    n2 = n2 - 1
                    s = s + 2
                    lkh = 0
                ElseIf n1 > 0 And n2 = 0 Then
                    '相当于中断结清
                    n1 = n1 - 1
                    zd = True
                    lkh = 1
                Else
                    If CLng(arr1(n1)) > CLng(arr2(n2)) Then
                        n1 = n1 - 1
                        zd = True
                    Else
                        n2 = n2 - 1
                        s = s + 2
                    End If
                End If
            Else
                '到这里就是中断
                zd = True
            End If
        End If
        
        If zd = True Or i = Len(rule) Then
            '结清
            If s > 0 Then
                
                If (i = Len(rule) And DoCompare_AddMyComplexExplain_RoundRule_IsNumber(Mid(rule, i, 1))) Or _
                    (i = Len(rule) And Mid(rule, i, 1) = ")" And lkh = 0) Then
                    If relust = "" Then
                        relust = CStr(i - s + 1) & "," & CStr(s)
                    Else
                        relust = relust & "," & CStr(i - s + 1) & "," & CStr(s)
                    End If
                Else
                    If relust = "" Then
                        relust = CStr(i - s) & "," & CStr(s)
                    Else
                        relust = relust & "," & CStr(i - s) & "," & CStr(s)
                    End If
                End If
                s = 0
            End If
        End If
        '下一个循环
        i = i + 1
nextI:

    Loop
    'Debug.Print relust
    '拆分出来了，下来就赋值
    Dim arr As Variant, j As Long
    arr = Split(relust, ",")
    i = 1
    j = 0
    rel = ""
    Do While j <= UBound(arr) Or i <= Len(rule)
        If j > UBound(arr) Then
            'j > 的话，说明剩下的都是
            rel = rel & Mid(rule, i, Len(rule) - i + 1)
            i = Len(rule)
            Exit Do
        ElseIf CLng(arr(j)) > i Then
            rel = rel & Mid(rule, i, CLng(arr(j)) - i)
            i = CLng(arr(j))
        Else
            rel = rel & "round(" & Mid(rule, CLng(arr(j)), CLng(arr(j + 1))) & "," & CStr(LEN_ROUND) & ")"
            i = CLng(arr(j)) + CLng(arr(j + 1))
            j = j + 2
        End If
        If i > Len(rule) Then Exit Do
    Loop
    DoCompare_AddMyComplexExplain_RoundRule = rel
Exit Function

errhandle:
    DoCompare_AddMyComplexExplain_RoundRule = strbk
End Function

'
'   获取关键字
'
Private Function DoCompare_AddMyComplexExplain_RoundRule_GetKeyType(rule As String, i As Long) As String
    Dim str As String, j As Long
    
    For j = 0 To UBound(arrComplexeKeys) Step 1
        str = CStr(arrComplexeKeys(j))
        
        If Len(rule) - i + 1 >= Len(str) And Mid(rule, i, Len(str)) = str Then
            DoCompare_AddMyComplexExplain_RoundRule_GetKeyType = str
            Exit Function
        End If
    Next j
    
    DoCompare_AddMyComplexExplain_RoundRule_GetKeyType = ""
End Function

'
'   数值计算
'
Private Function DoCompare_AddMyComplexExplain_RoundRule_IsNumber(str As String) As Boolean
    If str = "0" Or str = "1" Or str = "2" Or str = "3" Or str = "4" Or str = "5" Or str = "6" _
        Or str = "7" Or str = "8" Or str = "9" Or str = "." Or str = "+" Or str = "-" Or str = "*" Or str = "/" Then
        DoCompare_AddMyComplexExplain_RoundRule_IsNumber = True
    Else
        DoCompare_AddMyComplexExplain_RoundRule_IsNumber = False
    End If
End Function



Sub testreg()
    Dim s As String
    Dim reg As Object, mc As Object, m As Object, data_type As String, t_val As Variant
    Set reg = CreateObject("VBSCRIPT.REGEXP")
    With reg
        .Global = True: .ignorecase = False: .Pattern = "\[.*?\]"
    End With
    s = "[asdsd][fad]"
    Set mc = reg.Execute(s)
    For Each m In mc
        Debug.Print Mid(CStr(m.value), 2, Len(CStr(m.value)) - 2)
    Next
End Sub

'
'   初始化复杂指标表
'
Public Function InitComplexRef() As String
    Dim ws As Worksheet, i%, j%, k%, arr As Variant, arrT As Variant, code As String, str As String
    Dim reg As Object, mc As Object, m As Object, rule As String, rn As Range, errRuleNum As Long, rightRule As Boolean, ruleNum As Long, jyNum As Long
    Dim dtemp As Object
    Dim drule_col As Object
    
    glo_thd = ycV(gloCft.GetConfig("thd值"))
    arrComplexeKeys = Split(gloCft.GetConfig("系统导数比较-复杂校验-公式精度处理关键字数组"), ",")
    
    Set dComplexRef = Nothing
    Set dComplexRef = CreateObject("scripting.dictionary")
    Set drule_col = CreateObject("scripting.dictionary")

    Set reg = CreateObject("VBSCRIPT.REGEXP")
    dComplexRef.RemoveAll
    ComplexThd = 0  '误差项
    '单位转换
    ComplexThd = ycV(ThisWorkbook.Worksheets("更新日志").Range("C1").value)
    'ComplexThd = ComplexThd * (GetUnitVal("元") / GetUnitVal(DesUnit))
    
    Set ws = ThisWorkbook.Worksheets("公式校验")
    'Call gloEt.ShowData(ws)
    errRuleNum = 0
    ruleNum = 0

    Set rn = Nothing

    ws.Cells.Interior.colorindex = 0
    arr = ws.Range("A1").CurrentRegion.value   'ws.UsedRange.Value


    For i = 2 To UBound(arr) Step 1
        '判断是否禁用 - 空则标识启用，非空则表示禁用
        rightRule = True
        Set dtemp = CreateObject("scripting.dictionary")
        For j = 1 To UBound(arr, 2)
            dtemp(CStr(arr(1, j))) = CStr(arr(i, j))
        Next
        If (CStr(dtemp("是否启用")) = "" Or VarType(dtemp("是否启用")) = vbEmpty) And CStr(dtemp("校验规则")) <> "" Then
            If InStr(glo_data_pc, CStr(dtemp("适用批次"))) > 0 Then
                dtemp("校验规则") = InitComplexRef_ClearRule(CStr(dtemp("校验规则")))
                If CStr(dtemp("校验规则")) = "" Then
                    rightRule = False
                Else
                    dtemp("计算值公式") = InitComplexRef_ClearRule(CStr(dtemp("计算值公式")))
                    If Round(ycV(dtemp("thd值")), 10) = 0 Then
                        dtemp("thd值") = glo_thd
                    End If
                    If InitComplexRef_CheckTaget(dtemp("校验规则")) = False Then
                        rightRule = False
                    End If
                    If CStr(dtemp("取反标识")) = "" Then
                        dtemp("取反标识") = False
                    Else
                        dtemp("取反标识") = True
                    End If
                    If CStr(dtemp("计算值公式")) <> "" Then
                        If InitComplexRef_CheckTaget(dtemp("计算值公式")) = False Then
                            rightRule = False
                        End If
                    End If
                End If
            
                If rightRule = True Then
                    Set dComplexRef(i) = dtemp
                    ruleNum = ruleNum + 1
                Else
                    errRuleNum = errRuleNum + 1
                    If rn Is Nothing Then
                        Set rn = ws.Range(ws.Cells(i, 1), ws.Cells(i, UBound(arr, 2)))
                    Else
                        Set rn = Application.Union(rn, ws.Range(ws.Cells(i, 1), ws.Cells(i, UBound(arr, 2))))
                    End If
                End If
            End If
        End If
        Set dtemp = Nothing
    Next i
    
    If rn Is Nothing Then
    
    Else
        rn.Interior.colorindex = 15
        Call gloDe.InsertNewDebug("公式校验初始化", "-", ws.name, "初始化", "读取有效公式校验规则" & CStr(ruleNum) & "条，其中，未读取的错误校验规则" & CStr(errRuleNum) & "条。")
    End If
End Function

Private Function InitComplexRef_CheckTaget(rule As String) As Boolean
    
    InitComplexRef_CheckTaget = True
End Function

'
'   初始化校验规则,检查括号是否匹配
'
Private Function InitComplexRef_ClearRule(rule As String) As String
    Dim clearRule As String
    
    clearRule = replace(rule, "【", "[")
    clearRule = replace(clearRule, "】", "]")
    clearRule = replace(clearRule, "（", "(")
    clearRule = replace(clearRule, "）", ")")
    clearRule = replace(clearRule, "｛", "{")
    clearRule = replace(clearRule, "｝", "}")
    clearRule = replace(clearRule, "，", ",")
    clearRule = replace(clearRule, " ", " ")
    clearRule = replace(clearRule, "　", " ")
    
    
    If Len(replace(clearRule, "[", "")) <> Len(replace(clearRule, "]", "")) Or _
        Len(replace(clearRule, "(", "")) <> Len(replace(clearRule, ")", "")) Or _
        Len(replace(clearRule, "{", "")) <> Len(replace(clearRule, "}", "")) Then
        InitComplexRef_ClearRule = ""
    Else
        If Len(replace(clearRule, "[", "")) = Len(clearRule) Then
            InitComplexRef_ClearRule = ""
        Else
            InitComplexRef_ClearRule = clearRule
        End If
    End If
End Function



