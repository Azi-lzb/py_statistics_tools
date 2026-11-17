Attribute VB_Name = "C规则配置校验"
Private start_row As Long   ' 数据起始行
Private start_col As Long   ' 数据起始列
Private end_row As Long    ' 数据结束行
Private end_col As Long    ' 数据结束列
Private col_row As Integer ' 列编号行位置
Private t_name As String    ' 表格名称
Private cur_ws_data As Variant  ' 当张表数据
Private org_name As String
Private org_code As String
Private date_s As Variant
Private cur_drule As Object
Private d_cur_col As Object

Public Function g_d(org_name As String, target_code) As Variant
    If dCurDataRef.exists(org_name & "_" & target_code) Then
        get_d = dCurDataRef(org_name & "_" & target_code)
    Else
        get_d = ""
    End If
End Function

Private Function g_rule(Optional f_rel As String = "错误") As Object
    Dim drel_temp As Object, k As Variant, v As Variant, i As Integer
    k = cur_drule.keys
    v = cur_drule.items
    Set drel_temp = CreateObject("scripting.dictionary")
    For i = 0 To UBound(k)
        drel_temp(k(i)) = v(i)
    Next
    
    drel_temp("机构代码") = org_code
    drel_temp("是否执行异常") = "否": drel_temp("异常说明") = "": drel_temp("校验结果") = f_rel
        
    Set g_rule = drel_temp
End Function

'
'   便于取数
'
Private Function g_v(i As Long, col_s As String) As Double
    On Error Resume Next
    Dim v As Double
    v = 0
    If d_cur_col.exists(col_s) Then
        v = ycV(cur_ws_data(i, d_cur_col(col_s)))
    Else
        v = 0
    End If
    g_v = v
End Function

Private Function g_iv(i As Long, j As Integer) As Double
    On Error Resume Next
    Dim v As Double
    v = 0
    v = ycV(cur_ws_data(i, j))
    g_iv = v
End Function

Private Function g_is(i As Long, j As Integer) As String
    On Error Resume Next
    Dim v As String
    v = 0
    v = CStr(cur_ws_data(i, j))
    g_is = XCTR(v)
End Function

Private Function g_s(i As Long, col_s As String) As String
    On Error Resume Next
    Dim v As String
    v = 0
    If d_cur_col.exists(col_s) Then
        v = CStr(cur_ws_data(i, d_cur_col(col_s)))
    Else
        v = 0
    End If
    g_s = XCTR(v)
End Function

Private Function XCTR(str As Variant) As String
    Dim s As String
    s = replace(Trim(CStr(str)), "'", "")
    s = replace(s, "--", "")
    
    XCTR = s
End Function
    

'
'   分组执行
'
Public Sub CheckData_RuleCz()
    On Error Resume Next
    Dim k As Variant, i As Integer
    For Each glo_ws In glo_wb.Worksheets
        If glo_cz_table.exists(glo_ws.name) Then
            If dRuleRef.exists(glo_ws.name) Then
                k = dRuleRef(glo_ws.name).keys
                
                t_name = glo_ws.name
                start_row = CLng(ycV(Find_Cz(glo_cz_table, t_name, "数据起始行")))
                end_row = Get_Last_Row(glo_ws)
                col_row = CLng(ycV(Find_Cz(glo_cz_table, t_name, "列编号行位置")))
                start_col = CLng(ycV(Find_Cz(glo_cz_table, t_name, "数据起始列")))
                end_col = Get_Last_Col(glo_ws, col_row)
                If start_row = 0 Or col_row = 0 Or start_col = 0 Or end_row < start_row Or end_col < start_col Then
                    
                    Exit Sub
                End If
                
                org_name = replace(CStr(glo_ws.Range(Find_Cz(glo_cz_table, t_name, "机构名称位置")).value), "填报机构名称：", "")
                org_code = replace(CStr(glo_ws.Range(Find_Cz(glo_cz_table, t_name, "社会信用代码位置")).value), "填报机构统一社会信用代码：", "")
                date_s = ycD(glo_ws.Range(Find_Cz(glo_cz_table, t_name, "数据日期位置")).value)
                
                ' 正式开始读取
                If glo_aData.exists(t_name) Then
                    cur_ws_data = glo_aData(t_name)
                Else
                    cur_ws_data = gloEt.GetUsedRange(glo_ws).value
                End If
                
                Set d_cur_col = CreateObject("scripting.dictionary")
                For i = start_col To end_col
                    d_cur_col(CStr(cur_ws_data(start_col, i))) = i
                Next
                
                For i = 0 To UBound(k)
                    Set cur_drule = dRuleRef(glo_ws.name)(CStr(k(i)))
                    If InStr(glo_data_pc, CStr(cur_drule("启用批次"))) > 0 Then
                        Set d_aRel = CreateObject("scripting.dictionary")
                        Application.Run CStr(k(i))
                        If err.Number <> 0 Then
                            Debug.Print "函数(" & CStr(k(i)) & ")不存在？"
                            err.Clear
                        End If
                    Set cur_drule = Nothing
                    End If
                    If i Mod 3 = 0 Then Call gloCft.ShowStatusMessage_Up("正在进行规则校验，执行到表[" & glo_ws.name & "]" & CStr(k(i)) & "校验。")
                Next
                
                Erase cur_ws_data
                Set d_cur_col = Nothing
            End If
        Else
            ' 输出不存在
        End If
    Next
    
    If dRuleRef.exists("通用") Then
        k = dRuleRef("通用").keys
        
        For i = 0 To UBound(k)
            Set cur_drule = dRuleRef("通用")(CStr(k(i)))
            If InStr(glo_data_pc, CStr(cur_drule("启用批次"))) > 0 Then
                Set d_aRel = CreateObject("scripting.dictionary")
                Application.Run CStr(k(i))
                If err.Number <> 0 Then
                    Debug.Print "函数(" & CStr(k(i)) & ")不存在？"
                    err.Clear
                End If
            Set cur_drule = Nothing
            End If
             Call gloCft.ShowStatusMessage_Up("正在进行通用规则校验，执行到" & CStr(k(i)) & "校验。")
        Next
        Erase cur_ws_data
        Set d_cur_col = Nothing
    End If
End Sub

'
'   下面全部都是校验
'
Private Sub S0100_Rule_001()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1)

        
        v = g_v(i, "AE") - (g_v(i, "G") + g_v(i, "AB"))
        drel_temp("计算值") = v
        If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub


Private Sub S0100_Rule_002()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1)

        
        ' G=H+N+O+V+W+X+Y+Z+AA
        v = g_v(i, "G") - (g_v(i, "H") + g_v(i, "N") + g_v(i, "O") + g_v(i, "V") + g_v(i, "W") + g_v(i, "X") + g_v(i, "Y") + g_v(i, "Z") + g_v(i, "AA"))
        drel_temp("计算值") = v
        If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0100_Rule_003()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1)

        
        ' H≥I+K+L+M
        v = g_v(i, "H") - (g_v(i, "I") + g_v(i, "K") + g_v(i, "L") + g_v(i, "M"))
        drel_temp("计算值") = v
        If (Round(v, LEN_ROUND)) >= 0 Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0100_Rule_004()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1)
        
        ' I≥J
        v = g_v(i, "I") - g_v(i, "J")
        drel_temp("计算值") = v
        If (Round(v, LEN_ROUND)) >= 0 Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

' O=P+S+T
Private Sub S0100_Rule_005()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1)

        
        ' O=P+S+T
        v = g_v(i, "O") - (g_v(i, "P") + g_v(i, "S") + g_v(i, "T"))
        drel_temp("计算值") = v
        If Abs((Round(v, LEN_ROUND))) <= glo_thd Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0100_Rule_006()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1)

        ' P=Q+R
        v = g_v(i, "P") - (g_v(i, "Q") + g_v(i, "R"))
        drel_temp("计算值") = v
        If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

' T≥U
Private Sub S0100_Rule_007()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1)

        
        ' T≥U
        v = g_v(i, "T") - (g_v(i, "U"))
        drel_temp("计算值") = v
        If (Round(v, LEN_ROUND)) >= 0 Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

' AB=AC+AD
Private Sub S0100_Rule_008()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1)
        
        ' AB=AC+AD
        v = g_v(i, "AB") - (g_v(i, "AC") + g_v(i, "AD"))
        drel_temp("计算值") = v
        If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0100_Rule_009()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long, j As Integer, k As Integer
    Dim v2 As Double, is_10 As Boolean
    Set d_temp = cur_drule
    
    For i = start_row To end_row - 1 Step 11
        is_10 = False
        
        For j = 9 To end_col
            Set drel_temp = g_rule("错误")
            
            '可能变的地方
            drel_temp("校验描述") = ">=10户时，汇总行金额≥已列示10户法人机构明细求和。" 'CStr(d_temp("校验描述"))
            drel_temp("校验规则") = "汇总行数值≥下属10户法人机构明细行同列求和。"
            drel_temp("触发指标") = g_is(i, 1) & "_" & g_is(CLng(col_row), j)
            v2 = 0
            
            ' 汇总子项数据
            For k = 1 To 10
                v2 = v2 + g_iv(i + k, j)
                If g_s(i + k, "A") = "" Then
                    If k < 10 Then
                        is_10 = True
                        drel_temp("校验描述") = "一级交易对手下属法人机构<10户时，汇总行金额等于下属法人机构明细行求和。"
                        drel_temp("校验规则") = "汇总行数值=下属全部法人机构明细行同列求和。"
                        Exit For
                    End If
                End If
                v = g_iv(i, j) - v2
                drel_temp("计算值") = v
                If is_10 = True Then
                    If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
                        drel_temp("校验结果") = "正确"
                    End If
                Else
                    If Round(v, LEN_ROUND) >= 0 Then
                        drel_temp("校验结果") = "正确"
                    End If
                End If
            Next
            If err.Number <> 0 Then
                drel_temp("是否执行异常") = "是"
                drel_temp("异常说明") = err.Description
                err.Clear
            End If
            If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
                If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                    Set glo_rel_rule(glo_rel_rule_num) = drel_temp
                End If
            End If
            glo_rel_rule_num = glo_rel_rule_num + 1
            Set drel_temp = Nothing
        Next
        
    Next
End Sub

'
'
Private Sub S0100_Rule_010()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long, j As Integer, k As Integer
    Dim v2 As Double, is_50 As Boolean
    Set d_temp = cur_drule
    
    is_50 = False
    For i = start_row To end_row - 1 Step 11
        If g_s(i, "A") = "" Then
            If i / 11 < 50 Then
                is_50 = True
                Exit For
            End If
        End If
    Next
    
    For j = 9 To end_col
        v = g_iv(end_row, j)
        v2 = 0
        Set drel_temp = g_rule("错误")
        drel_temp("校验描述") = ">=50户时，总合计≥已列示50户一级交易对手汇总行求和。" 'CStr(d_temp("校验描述"))
        If is_50 = True Then drel_temp("校验描述") = "全部一级交易对手<50户时，总合计等于全部一级交易对手汇总行求和。"
        
        drel_temp("校验规则") = "100000列数值≥101000+102000+…+150000同列数值。" 'CStr(d_temp("校验描述"))
        If is_50 = True Then drel_temp("校验规则") = "100000列数值=101000+102000+…+150000同列数值。"
        drel_temp("触发指标") = g_is(end_row, 1) & "_" & g_is(CLng(col_row), j)
        
        
        For i = start_row To end_row - 1 Step 11
            v2 = v2 + g_iv(i, j)
        Next
        
            

        v = v - v2
        drel_temp("计算值") = v
        If is_50 = True Then
            If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
                drel_temp("校验结果") = "正确"
            End If
        Else
            If Round(v, LEN_ROUND) >= 0 Then
                drel_temp("校验结果") = "正确"
            End If
        End If

        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
        
    Next
End Sub

'
' 全部一级交易对手按余额合计从大到小降序排列
'
Private Sub S0100_Rule_011()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long, j As Integer, k As Integer
    Dim v2 As Double, is_50 As Boolean
    Set d_temp = cur_drule
    

    v = g_v(end_row, "AE")
    v2 = 0
    Set drel_temp = g_rule("正确")
    drel_temp("触发指标") = "100000_所有交易对手合计_余额合计"
    
    
    v = g_v(start_row, "AE")
    For i = start_row To end_row - 1 Step 11
        If i <> start_row Then
            v2 = g_v(i, "AE")
            If Round(v - v2, LEN_ROUND) < 0 Then
                drel_temp("触发指标") = g_is(i - 11, 1) & " < " & g_is(i, 1)
                drel_temp("计算值") = Round(v - v2, LEN_ROUND)
                drel_temp("校验结果") = "错误"
                Exit For
            Else
                v = v2
            End If
        End If
    Next
    

    If err.Number <> 0 Then
        drel_temp("是否执行异常") = "是"
        drel_temp("异常说明") = err.Description
        err.Clear
    End If
    If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
        If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
            Set glo_rel_rule(glo_rel_rule_num) = drel_temp
        End If
    End If
    glo_rel_rule_num = glo_rel_rule_num + 1
    Set drel_temp = Nothing
End Sub

Private Sub S0100_Rule_012()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long, j As Integer, k As Integer
    Dim v2 As Double, is_50 As Boolean
    Set d_temp = cur_drule
    

    v = g_v(end_row, "AE")
    v2 = 0

    For i = start_row To end_row - 1 Step 11
        Set drel_temp = g_rule("正确")
        drel_temp("触发指标") = g_is(i, 1) & "_余额合计"
        For j = 1 To 10
            If j = 1 Then
                v = g_v(i + j, "AE")
            Else
                v2 = g_v(i + j, "AE")
                If Round(v - v2, LEN_ROUND) < 0 Then
                    drel_temp("触发指标") = g_is(i + j - 1, 1) & " < " & g_is(i + j, 1)
                    drel_temp("计算值") = Round(v - v2, LEN_ROUND)
                    drel_temp("校验结果") = "错误"
                    Exit For
                Else
                    v = v2
                End If
            End If
        Next
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub



Private Sub S0100_Rule_013()
    On Error Resume Next
    
    Dim r As String, drel_temp As Object, d_temp As Object, i As Long, j As Integer, k As Integer
    Set d_temp = cur_drule
    

    For i = start_row To end_row - 1
        Set drel_temp = g_rule("正确")
        drel_temp("触发指标") = g_is(i, 1) & "_" & g_is(i, 4) & "_" & g_is(i, 3)
        If g_is(i, 4) = "" Then
        
        Else
            r = ValidateUSCC(glo_ws.Cells(i, 4))
            If r = "" Then
            
            Else
                drel_temp("校验描述") = CStr(drel_temp("校验描述")) & r
                drel_temp("校验结果") = "错误"
            End If
            
            If err.Number <> 0 Then
                drel_temp("是否执行异常") = "是"
                drel_temp("异常说明") = err.Description
                err.Clear
            End If
            If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
                If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                    Set glo_rel_rule(glo_rel_rule_num) = drel_temp
                End If
            End If
            glo_rel_rule_num = glo_rel_rule_num + 1
        End If
        
        
        Set drel_temp = Nothing
    Next
End Sub


Function ValidateUSCC(cell As Range) As String
    Dim val As Variant
    
    If IsEmpty(cell) Then
        ValidateUSCC = "单元格为空"
        Exit Function
    End If
    
    val = cell.value
    
    '检查是否为数字类型（被数值化）
    'If IsNumeric(val) Then
    '    ValidateUSCC = "已被数值化，请改为文本格式"
    '    Exit Function
    'End If
    
    '检查是否为文本格式
    If cell.NumberFormat <> "@" Then
        ValidateUSCC = "不是文本格式单元格"
        Exit Function
    End If
    
    ValidateUSCC = ""
End Function

Private Function InitBaseInfo(ws_name As String) As Boolean
    On Error Resume Next
    t_name = ws_name
    Set glo_ws = glo_wb.Worksheets(t_name)
    start_row = CLng(ycV(Find_Cz(glo_cz_table, t_name, "数据起始行")))
    end_row = Get_Last_Row(glo_wb.Worksheets(t_name))
    col_row = CLng(ycV(Find_Cz(glo_cz_table, t_name, "列编号行位置")))
    start_col = CLng(ycV(Find_Cz(glo_cz_table, t_name, "数据起始列")))
    end_col = Get_Last_Col(glo_wb.Worksheets(t_name), col_row)
    
    
    If start_row = 0 Or col_row = 0 Or start_col = 0 Or end_row < start_row Or end_col < start_col Then
        InitBaseInfo = False
    Else
        org_name = replace(CStr(glo_ws.Range(Find_Cz(glo_cz_table, t_name, "机构名称位置")).value), "填报机构名称：", "")
        org_code = replace(CStr(glo_ws.Range(Find_Cz(glo_cz_table, t_name, "社会信用代码位置")).value), "填报机构统一社会信用代码：", "")
        date_s = ycD(glo_ws.Range(Find_Cz(glo_cz_table, t_name, "数据日期位置")).value)
        
        ' 正式开始读取
        If glo_aData.exists(t_name) Then
            cur_ws_data = glo_aData(t_name)
        Else
            cur_ws_data = gloEt.GetUsedRange(glo_ws).value
        End If
        
        Set d_cur_col = Nothing
        Set d_cur_col = CreateObject("scripting.dictionary")
        For i = start_col To end_col
            d_cur_col(CStr(cur_ws_data(start_col, i))) = i
        Next
        InitBaseInfo = True
    End If
    
    If err.Number <> 0 Then
        Debug.Print "InitBaseInfo Error - " & err.Description
        InitBaseInfo = False
        err.Clear
    End If
End Function

Private Sub Sample_Rule_001()
    On Error Resume Next
    
    '判断至少需要存在两张表
    
    '读取每个表所有交易对手名称、代码，以及相关信息。如果是一致的，那我干脆直接就拼接在一起就好了
    Dim t As String, arrT As Variant, ti As Integer
    Dim r As String, drel_temp As Object, d_temp As Object, i As Long, j As Integer, k As Integer
    Dim d_key As String, is_Right As Boolean
    
    is_Right = True
    Set d_temp = cur_drule
    t = "S0100_F_26,S0100_J_26,S0200_F_26,S0200_J_26,S0300_F_26,S0300_J_26" ' ,S0400_F_26"
    arrT = Split(t, ",")
    If UBound(arrT) < 1 Then
        Exit Sub
    End If
    
    Dim d_name As Object, d_code As Object
    Set d_name = CreateObject("scripting.dictionary")
    Set d_code = CreateObject("scripting.dictionary")
    
    
    For ti = 0 To UBound(arrT)
        If glo_aData.exists(CStr(arrT(ti))) Then
            
            If InitBaseInfo(CStr(arrT(ti))) = True Then
                For i = start_row To end_row - 1
                    If g_is(i, 4) = "" Then
                    
                    Else
                        d_key = t_name & "|" & g_is(i, 3) & "," & g_is(i, 4) & "," & g_is(i, 5) & "," & g_is(i, 7) & "," & g_is(i, 8)
                        If d_name.exists(g_is(i, 3)) Then
                            If ycRight(d_key, "|", 1, 0) = ycRight(CStr(d_name(g_is(i, 3))), "|", 1, 0) Then
                            
                            Else
                                Set drel_temp = g_rule("错误")
                                drel_temp("校验表单") = "表1-3"
                                drel_temp("校验描述") = CStr(drel_temp("校验描述")) & "-按名称"
                                drel_temp("触发指标") = g_is(i, 1) & "_" & g_is(i, 4) & "_" & g_is(i, 3)
                                drel_temp("计算值") = CStr(d_name(g_is(i, 3))) & "<>" & d_key
                                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
                                glo_rel_rule_num = glo_rel_rule_num + 1
                                Set drel_temp = Nothing
                                is_Right = False
                            End If
                        Else
                            d_name(g_is(i, 3)) = d_key
                        End If
                        
                        If d_code.exists(g_is(i, 4)) Then
                            If ycRight(d_key, "|", 1, 0) = ycRight(CStr(d_name(g_is(i, 4))), "|", 1, 0) Then
                            
                            Else
                                Set drel_temp = g_rule("错误")
                                drel_temp("校验表单") = "表1-3"
                                drel_temp("校验描述") = CStr(drel_temp("校验描述")) & "-按代码"
                                drel_temp("触发指标") = g_is(i, 1) & "_" & g_is(i, 4) & "_" & g_is(i, 3)
                                drel_temp("计算值") = CStr(d_name(g_is(i, 4))) & "<>" & d_key
                                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
                                glo_rel_rule_num = glo_rel_rule_num + 1
                                Set drel_temp = Nothing
                                is_Right = False
                            End If
                        Else
                            d_name(g_is(i, 4)) = d_key
                        End If
                    End If
                Next
            End If
        End If
    Next
    
    If is_Right = True Then
        Set drel_temp = g_rule("正确")
        drel_temp("校验表单") = "表1-1 - 表1-3"
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
                glo_rel_rule_num = glo_rel_rule_num + 1
            End If
        End If
        Set drel_temp = Nothing
    End If
        
End Sub

Private Sub Sample_Rule_003()
    On Error Resume Next
    
    '判断至少需要存在两张表
    
    '读取每个表所有交易对手名称、代码，以及相关信息。如果是一致的，那我干脆直接就拼接在一起就好了
    Dim t As String, arrT As Variant, ti As Integer
    Dim r As String, drel_temp As Object, d_temp As Object, i As Long, j As Integer, k As Integer
    Dim d_key As String, is_Right As Boolean
    
    is_Right = True
    Set d_temp = cur_drule
    t = "S0100_F_26,S0100_J_26,S0200_F_26,S0200_J_26,S0300_F_26,S0300_J_26,S0400_F_26" '
    arrT = Split(t, ",")
    If UBound(arrT) < 1 Then
        Exit Sub
    End If
    
    Dim d_name As Object, d_code As Object
    Set d_name = CreateObject("scripting.dictionary")
    Set d_code = CreateObject("scripting.dictionary")
    
    
    For ti = 0 To UBound(arrT)
        If glo_aData.exists(CStr(arrT(ti))) Then
            
            If InitBaseInfo(CStr(arrT(ti))) = True Then
                For i = start_row To end_row - 1
                    If g_is(i, 4) = "" Then
                    
                    Else
                        d_key = t_name & "|" & g_is(i, 3) & "," & g_is(i, 4)
                        If d_name.exists(g_is(i, 3)) Then
                            If ycRight(d_key, "|", 1, 0) = ycRight(CStr(d_name(g_is(i, 3))), "|", 1, 0) Then
                            
                            Else
                                Set drel_temp = g_rule("错误")
                                drel_temp("校验表单") = "表1-3"
                                drel_temp("校验描述") = CStr(drel_temp("校验描述")) & "-按名称"
                                drel_temp("触发指标") = g_is(i, 1) & "_" & g_is(i, 4) & "_" & g_is(i, 3)
                                drel_temp("计算值") = CStr(d_name(g_is(i, 3))) & "<>" & d_key
                                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
                                glo_rel_rule_num = glo_rel_rule_num + 1
                                Set drel_temp = Nothing
                                is_Right = False
                            End If
                        Else
                            d_name(g_is(i, 3)) = d_key
                        End If
                        
                        If d_code.exists(g_is(i, 4)) Then
                            If ycRight(d_key, "|", 1, 0) = ycRight(CStr(d_name(g_is(i, 4))), "|", 1, 0) Then
                            
                            Else
                                Set drel_temp = g_rule("错误")
                                drel_temp("校验表单") = "表1-3"
                                drel_temp("校验描述") = CStr(drel_temp("校验描述")) & "-按代码"
                                drel_temp("触发指标") = g_is(i, 1) & "_" & g_is(i, 4) & "_" & g_is(i, 3)
                                drel_temp("计算值") = CStr(d_name(g_is(i, 4))) & "<>" & d_key
                                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
                                glo_rel_rule_num = glo_rel_rule_num + 1
                                Set drel_temp = Nothing
                                is_Right = False
                            End If
                        Else
                            d_name(g_is(i, 4)) = d_key
                        End If
                    End If
                Next
            End If
        End If
    Next
    
    If is_Right = True Then
        Set drel_temp = g_rule("正确")
        drel_temp("校验表单") = "表1-1 - 表1-4"
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
                glo_rel_rule_num = glo_rel_rule_num + 1
            End If
        End If
        Set drel_temp = Nothing
    End If
        
End Sub

Private Sub Sample_Rule_002()
    On Error Resume Next
    
    On Error Resume Next
    
    '判断至少需要存在两张表
    
    '读取每个表所有交易对手名称、代码，以及相关信息。如果是一致的，那我干脆直接就拼接在一起就好了
    Dim t As String, arrT As Variant, ti As Integer
    Dim r As String, drel_temp As Object, d_temp As Object, i As Long, j As Integer, k As Integer
    
    
    Set d_temp = cur_drule
    t = "S0100_F_26,S0100_J_26,S0200_F_26,S0200_J_26,S0300_F_26,S0300_J_26,S0400_F_26"
    arrT = Split(t, ",")
    If UBound(arrT) < 1 Then
        Exit Sub
    End If
    
    For ti = 0 To UBound(arrT)
        If glo_aData.exists(CStr(arrT(ti))) Then
            If InitBaseInfo(CStr(arrT(ti))) = True Then
                is_Right = True
                For i = start_row To end_row - 1
                    Set drel_temp = g_rule("正确")
                    drel_temp("校验表单") = t_name
                    drel_temp("触发指标") = g_is(i, 1) & "_" & g_is(i, 4) & "_" & g_is(i, 3)
                    If g_is(i, 4) = "" Then
                    
                    Else
                        r = ValidateUSCC(glo_ws.Cells(i, 4))
                        If r = "" Then
                        
                        Else
                            drel_temp("校验描述") = CStr(drel_temp("校验描述")) & r
                            drel_temp("校验结果") = "错误"
                            is_Right = False
                        End If
                        
                        If err.Number <> 0 Then
                            drel_temp("校验结果") = "错误"
                            drel_temp("是否执行异常") = "是"
                            drel_temp("异常说明") = err.Description
                            err.Clear
                        End If
                        If CStr(drel_temp("校验结果")) = "错误" Then
                            Set glo_rel_rule(glo_rel_rule_num) = drel_temp
                            glo_rel_rule_num = glo_rel_rule_num + 1
                        End If
                        
                    End If
                    
                    Set drel_temp = Nothing
                Next
                
                If is_Right = True Then
                    Set drel_temp = g_rule("正确")
                    drel_temp("校验表单") = t_name
                    drel_temp("触发指标") = "所有指标"
                    If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
                        If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                            Set glo_rel_rule(glo_rel_rule_num) = drel_temp
                        End If
                    End If
                    glo_rel_rule_num = glo_rel_rule_num + 1
                End If
            End If
        End If
    Next
End Sub


Private Sub S0200_Rule_001()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row - 1
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1)
        
        'G≥H+I+J+K+L+M+N+O
        v = g_v(i, "G") - (g_v(i, "H") + g_v(i, "I") + g_v(i, "J") + g_v(i, "K") + g_v(i, "L") + g_v(i, "M") + g_v(i, "N") + g_v(i, "O"))
        drel_temp("计算值") = v
        If (Round(v, LEN_ROUND)) >= 0 Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0200_Rule_002()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row - 1
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1)
        
        'Q≥R+S
        v = g_v(i, "Q") - (g_v(i, "R") + g_v(i, "S"))
        drel_temp("计算值") = v
        If (Round(v, LEN_ROUND)) >= 0 Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0200_Rule_003()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row - 1
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1)
        
        v = g_v(i, "T") - (g_v(i, "U") + g_v(i, "V"))
        drel_temp("计算值") = v
        If (Round(v, LEN_ROUND)) >= 0 Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0200_Rule_004()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row - 1
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1)
        
        v = g_v(i, "W") - (g_v(i, "X") + g_v(i, "Y"))
        drel_temp("计算值") = v
        If (Round(v, LEN_ROUND)) >= 0 Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0200_Rule_005()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row - 1
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1)
        
        v = g_v(i, "AA") - (g_v(i, "G") + g_v(i, "P") + g_v(i, "Q") + g_v(i, "T") + g_v(i, "W") + g_v(i, "Z"))
        drel_temp("计算值") = v
        If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0200_Rule_006()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row - 1
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1)
        
        v = g_v(i, "AA") - (g_v(i, "AB"))
        drel_temp("计算值") = v
        If (Round(v, LEN_ROUND)) >= 0 Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0200_Rule_007()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row - 1
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1)
        
        v = g_v(i, "AA") - 0
        drel_temp("计算值") = v
        If (Round(v, LEN_ROUND)) >= 0 Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0200_Rule_008()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row - 1
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1)
        
        v = g_v(i, "AD") - (g_v(i, "AA") + g_v(i, "AC"))
        drel_temp("计算值") = v
        If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0200_Rule_009()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long, j As Integer, k As Integer
    Dim v2 As Double, is_10 As Boolean
    Set d_temp = cur_drule
    
    For i = start_row To end_row - 5 Step 11
        is_10 = False
        
        For j = 10 To end_col
            Set drel_temp = g_rule("错误")
            
            '可能变的地方
            drel_temp("校验描述") = ">=10户时，汇总行金额≥已列示10户法人机构明细求和。" 'CStr(d_temp("校验描述"))
            drel_temp("校验规则") = "汇总行数值≥下属10户法人机构明细行同列求和。"
            drel_temp("触发指标") = g_is(i, 1) & "_" & g_is(CLng(col_row), j)
            v2 = 0
            
            ' 汇总子项数据
            For k = 1 To 10
                v2 = v2 + g_iv(i + k, j)
                If g_s(i + k, "A") = "" Then
                    If k < 10 Then
                        is_10 = True
                        drel_temp("校验描述") = "一级交易对手下属法人机构<10户时，汇总行金额等于下属法人机构明细行求和。"
                        drel_temp("校验规则") = "汇总行数值=下属全部法人机构明细行同列求和。"
                        Exit For
                    End If
                End If
                v = g_iv(i, j) - v2
                drel_temp("计算值") = v
                If is_10 = True Then
                    If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
                        drel_temp("校验结果") = "正确"
                    End If
                Else
                    If Round(v, LEN_ROUND) >= 0 Then
                        drel_temp("校验结果") = "正确"
                    End If
                End If
            Next
            If err.Number <> 0 Then
                drel_temp("是否执行异常") = "是"
                drel_temp("异常说明") = err.Description
                err.Clear
            End If
            If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
                If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                    Set glo_rel_rule(glo_rel_rule_num) = drel_temp
                End If
            End If
            glo_rel_rule_num = glo_rel_rule_num + 1
            Set drel_temp = Nothing
        Next
        
    Next
End Sub

'
'  全部一级交易对手≤50户时，总合计等于全部一级交易对手汇总行求和
'
Private Sub S0200_Rule_010()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long, j As Integer, k As Integer
    Dim v2 As Double, is_50 As Boolean
    Set d_temp = cur_drule
    
    is_50 = False
    For i = start_row To end_row - 5 Step 11
        If g_s(i, "A") = "" Then
            If i / 11 < 50 Then
                is_50 = True
                Exit For
            End If
        End If
    Next
    
    For j = 10 To end_col
        v = g_iv(end_row - 4, j)
        v2 = 0
        Set drel_temp = g_rule("错误")
        drel_temp("校验描述") = ">=50户时，总合计≥已列示50户一级交易对手汇总行求和。" 'CStr(d_temp("校验描述"))
        If is_50 = True Then drel_temp("校验描述") = "全部一级交易对手<50户时，总合计等于全部一级交易对手汇总行求和。"
        
        drel_temp("校验规则") = "200000列数值≥201000+202000+…+250000同列数值。" 'CStr(d_temp("校验描述"))
        If is_50 = True Then drel_temp("校验规则") = "200000列数值=201000+202000+…+250000同列数值。"
        drel_temp("触发指标") = g_is(end_row - 4, 1) & "_" & g_is(CLng(col_row), j)
        
        
        For i = start_row To end_row - 5 Step 11
            v2 = v2 + g_iv(i, j)
        Next
        
            

        v = v - v2
        drel_temp("计算值") = v
        If is_50 = True Then
            If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
                drel_temp("校验结果") = "正确"
            End If
        Else
            If Round(v, LEN_ROUND) >= 0 Then
                drel_temp("校验结果") = "正确"
            End If
        End If

        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
        
    Next
End Sub


Private Sub S0200_Rule_011()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long, j As Integer, k As Integer
    Dim v2 As Double, is_50 As Boolean
    Set d_temp = cur_drule
    
    v2 = 0
    Set drel_temp = g_rule("正确")
    drel_temp("触发指标") = "2a0000_前50大交易对手风险暴露小计_全部风险暴露总计"
    
    
    v = g_v(start_row, "AA")
    For i = start_row To end_row - 5 Step 11
        If i <> start_row Then
            v2 = g_v(i, "AA")
            If Round(v - v2, LEN_ROUND) < 0 Then
                drel_temp("触发指标") = g_is(i - 11, 1) & " < " & g_is(i, 1)
                drel_temp("计算值") = Round(v - v2, LEN_ROUND)
                drel_temp("校验结果") = "错误"
                Exit For
            Else
                v = v2
            End If
        End If
    Next
    

    If err.Number <> 0 Then
        drel_temp("是否执行异常") = "是"
        drel_temp("异常说明") = err.Description
        err.Clear
    End If
    If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
        If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
            Set glo_rel_rule(glo_rel_rule_num) = drel_temp
        End If
    End If
    glo_rel_rule_num = glo_rel_rule_num + 1
    Set drel_temp = Nothing
End Sub

Private Sub S0200_Rule_012()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long, j As Integer, k As Integer
    Dim v2 As Double, is_50 As Boolean
    Set d_temp = cur_drule
    

    v2 = 0

    For i = start_row To end_row - 5 Step 11
        Set drel_temp = g_rule("正确")
        drel_temp("触发指标") = g_is(i, 1) & "_全部风险暴露总计"
        For j = 1 To 10
            If j = 1 Then
                v = g_v(i + j, "AA")
            Else
                v2 = g_v(i + j, "AA")
                If Round(v - v2, LEN_ROUND) < 0 Then
                    drel_temp("触发指标") = g_is(i + j - 1, 1) & " < " & g_is(i + j, 1)
                    drel_temp("计算值") = Round(v - v2, LEN_ROUND)
                    drel_temp("校验结果") = "错误"
                    Exit For
                Else
                    v = v2
                End If
            End If
        Next
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0300_Rule_001()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1) & "_" & CStr(drel_temp("触发指标"))

        
        v = g_v(i, "G") - (g_v(i, "H") + g_v(i, "I") + g_v(i, "J") + g_v(i, "K") + g_v(i, "L") + g_v(i, "M") + g_v(i, "N") + g_v(i, "O"))
        drel_temp("计算值") = v
        If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0300_Rule_002()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1) & "_" & CStr(drel_temp("触发指标"))

        
        v = g_v(i, "Q") - (g_v(i, "G") + g_v(i, "P"))
        drel_temp("计算值") = v
        If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0300_Rule_003()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long, j As Integer, k As Integer
    Dim v2 As Double, is_10 As Boolean
    Set d_temp = cur_drule
    
    For i = start_row To end_row - 1 Step 11
        is_10 = False
        
        For j = 9 To end_col
            Set drel_temp = g_rule("错误")
            
            '可能变的地方
            drel_temp("校验描述") = ">=10户时，汇总行金额≥已列示10户法人机构明细求和。" 'CStr(d_temp("校验描述"))
            drel_temp("校验规则") = "汇总行数值≥下属10户法人机构明细行同列求和。"
            drel_temp("触发指标") = g_is(i, 1) & "_" & g_is(CLng(col_row), j)
            v2 = 0
            
            ' 汇总子项数据
            For k = 1 To 10
                v2 = v2 + g_iv(i + k, j)
                If g_s(i + k, "A") = "" Then
                    If k < 10 Then
                        is_10 = True
                        drel_temp("校验描述") = "一级交易对手下属法人机构<10户时，汇总行金额等于下属法人机构明细行求和。"
                        drel_temp("校验规则") = "汇总行数值=下属全部法人机构明细行同列求和。"
                        Exit For
                    End If
                End If
                v = g_iv(i, j) - v2
                drel_temp("计算值") = v
                If is_10 = True Then
                    If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
                        drel_temp("校验结果") = "正确"
                    End If
                Else
                    If Round(v, LEN_ROUND) >= 0 Then
                        drel_temp("校验结果") = "正确"
                    End If
                End If
            Next
            If err.Number <> 0 Then
                drel_temp("是否执行异常") = "是"
                drel_temp("异常说明") = err.Description
                err.Clear
            End If
            If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
                If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                    Set glo_rel_rule(glo_rel_rule_num) = drel_temp
                End If
            End If
            glo_rel_rule_num = glo_rel_rule_num + 1
            Set drel_temp = Nothing
        Next
        
    Next
End Sub

'
'
Private Sub S0300_Rule_004()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long, j As Integer, k As Integer
    Dim v2 As Double, is_50 As Boolean
    Set d_temp = cur_drule
    
    is_50 = False
    For i = start_row To end_row - 1 Step 11
        If g_s(i, "A") = "" Then
            If i / 11 < 50 Then
                is_50 = True
                Exit For
            End If
        End If
    Next
    
    For j = 9 To end_col
        v = g_iv(end_row, j)
        v2 = 0
        Set drel_temp = g_rule("错误")
        drel_temp("校验描述") = ">=50户时，总合计≥已列示50户一级交易对手汇总行求和。" 'CStr(d_temp("校验描述"))
        If is_50 = True Then drel_temp("校验描述") = "全部一级交易对手<50户时，总合计等于全部一级交易对手汇总行求和。"
        
        drel_temp("校验规则") = "300000列数值≥301000+302000+…+350000同列数值。" 'CStr(d_temp("校验描述"))
        If is_50 = True Then drel_temp("校验规则") = "300000列数值=301000+302000+…+350000同列数值。"
        drel_temp("触发指标") = g_is(end_row, 1) & "_" & g_is(CLng(col_row), j)
        
        
        For i = start_row To end_row - 1 Step 11
            v2 = v2 + g_iv(i, j)
        Next
        
        v = v - v2
        drel_temp("计算值") = v
        If is_50 = True Then
            If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
                drel_temp("校验结果") = "正确"
            End If
        Else
            If Round(v, LEN_ROUND) >= 0 Then
                drel_temp("校验结果") = "正确"
            End If
        End If

        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
                glo_rel_rule_num = glo_rel_rule_num + 1
            End If
        End If
        Set drel_temp = Nothing
        
    Next
End Sub


'
' 全部一级交易对手按余额合计从大到小降序排列
'
Private Sub S0300_Rule_005()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long, j As Integer, k As Integer
    Dim v2 As Double, is_50 As Boolean
    Set d_temp = cur_drule
    

    v2 = 0
    Set drel_temp = g_rule("正确")
    drel_temp("触发指标") = "300000_所有交易对手合计_融资余额合计"
    
    
    v = g_v(start_row, "Q")
    For i = start_row To end_row - 1 Step 11
        If i <> start_row Then
            v2 = g_v(i, "Q")
            If Round(v - v2, LEN_ROUND) < 0 Then
                drel_temp("触发指标") = g_is(i - 11, 1) & " < " & g_is(i, 1)
                drel_temp("计算值") = Round(v - v2, LEN_ROUND)
                drel_temp("校验结果") = "错误"
                Exit For
            Else
                v = v2
            End If
        End If
    Next
    

    If err.Number <> 0 Then
        drel_temp("是否执行异常") = "是"
        drel_temp("异常说明") = err.Description
        err.Clear
    End If
    If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
        If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
            Set glo_rel_rule(glo_rel_rule_num) = drel_temp
        End If
    End If
    glo_rel_rule_num = glo_rel_rule_num + 1
    Set drel_temp = Nothing
End Sub

Private Sub S0300_Rule_006()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long, j As Integer, k As Integer
    Dim v2 As Double, is_50 As Boolean
    Set d_temp = cur_drule
    
    v2 = 0

    For i = start_row To end_row - 1 Step 11
        Set drel_temp = g_rule("正确")
        drel_temp("触发指标") = g_is(i, 1) & "_融资余额合计"
        For j = 1 To 10
            If j = 1 Then
                v = g_v(i + j, "Q")
            Else
                v2 = g_v(i + j, "Q")
                If Round(v - v2, LEN_ROUND) < 0 Then
                    drel_temp("触发指标") = g_is(i + j - 1, 1) & " < " & g_is(i + j, 1)
                    drel_temp("计算值") = Round(v - v2, LEN_ROUND)
                    drel_temp("校验结果") = "错误"
                    Exit For
                Else
                    v = v2
                End If
            End If
        Next
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub


Private Sub S0400_Rule_001()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1) & "_" & CStr(drel_temp("触发指标"))

        
        v = g_v(i, "AA") - (g_v(i, "C") + g_v(i, "X"))
        drel_temp("计算值") = v
        If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0400_Rule_002()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1) & "_" & CStr(drel_temp("触发指标"))

        ' C=D+J+K+R+S+T+V+W
        v = g_v(i, "C") - (g_v(i, "D") + g_v(i, "J") + g_v(i, "K") + g_v(i, "R") + g_v(i, "S") + g_v(i, "T") + g_v(i, "V") + g_v(i, "W"))
        drel_temp("计算值") = v
        If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0400_Rule_003()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1) & "_" & CStr(drel_temp("触发指标"))

        ' D≥E+G+H+I
        v = g_v(i, "D") - (g_v(i, "E") + g_v(i, "G") + g_v(i, "H") + g_v(i, "I"))
        drel_temp("计算值") = v
        If (Round(v, LEN_ROUND)) >= 0 Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0400_Rule_004()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1) & "_" & CStr(drel_temp("触发指标"))


        v = g_v(i, "E") - (g_v(i, "F"))
        drel_temp("计算值") = v
        If (Round(v, LEN_ROUND)) >= 0 Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0400_Rule_005()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1) & "_" & CStr(drel_temp("触发指标"))

        
        v = g_v(i, "K") - (g_v(i, "L") + g_v(i, "O") + g_v(i, "P"))
        drel_temp("计算值") = v
        If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0400_Rule_006()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1) & "_" & CStr(drel_temp("触发指标"))

        
        v = g_v(i, "L") - (g_v(i, "M") + g_v(i, "N"))
        drel_temp("计算值") = v
        If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0400_Rule_007()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1) & "_" & CStr(drel_temp("触发指标"))


        v = g_v(i, "P") - (g_v(i, "Q"))
        drel_temp("计算值") = v
        If (Round(v, LEN_ROUND)) >= 0 Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0400_Rule_008()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1) & "_" & CStr(drel_temp("触发指标"))

        
        v = g_v(i, "X") - (g_v(i, "Y") + g_v(i, "Z"))
        drel_temp("计算值") = v
        If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0400_Rule_009()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long, j As Integer, k As Integer
    Dim v2 As Double, is_10 As Boolean
    Set d_temp = cur_drule
    
    For i = start_row To end_row - 2 Step 11
        is_10 = False
        
        For j = 5 To end_col
            Set drel_temp = g_rule("错误")
            
            '可能变的地方
            drel_temp("校验描述") = ">=10户时，汇总行金额≥已列示10户法人机构明细求和。" 'CStr(d_temp("校验描述"))
            drel_temp("校验规则") = "汇总行数值≥下属10户法人机构明细行同列求和。"
            drel_temp("触发指标") = g_is(i, 1) & "_" & g_is(CLng(col_row), j)
            v2 = 0
            
            ' 汇总子项数据
            For k = 1 To 10
                v2 = v2 + g_iv(i + k, j)
                If g_s(i + k, "A") = "" Then
                    If k < 10 Then
                        is_10 = True
                        drel_temp("校验描述") = "一级交易对手下属法人机构<10户时，汇总行金额等于下属法人机构明细行求和。"
                        drel_temp("校验规则") = "汇总行数值=下属全部法人机构明细行同列求和。"
                        Exit For
                    End If
                End If
                v = g_iv(i, j) - v2
                drel_temp("计算值") = v
                If is_10 = True Then
                    If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
                        drel_temp("校验结果") = "正确"
                    End If
                Else
                    If Round(v, LEN_ROUND) >= 0 Then
                        drel_temp("校验结果") = "正确"
                    End If
                End If
            Next
            If err.Number <> 0 Then
                drel_temp("是否执行异常") = "是"
                drel_temp("异常说明") = err.Description
                err.Clear
            End If
            If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
                If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                    Set glo_rel_rule(glo_rel_rule_num) = drel_temp
                End If
            End If
            glo_rel_rule_num = glo_rel_rule_num + 1
            Set drel_temp = Nothing
        Next
        
    Next
End Sub

Private Sub S0400_Rule_010()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long, j As Integer, k As Integer
    Dim v2 As Double, is_50 As Boolean
    Set d_temp = cur_drule
    
    is_50 = True
    
    For j = 5 To end_col
        v = g_iv(end_row - 1, j)
        v2 = 0
        Set drel_temp = g_rule("错误")
        drel_temp("触发指标") = g_is(end_row, 1) & "_" & g_is(CLng(col_row), j)
        
        
        For i = start_row To end_row - 2 Step 11
            v2 = v2 + g_iv(i, j)
        Next
        
        v = v - v2
        drel_temp("计算值") = v
        If is_50 = True Then
            If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
                drel_temp("校验结果") = "正确"
            End If
        Else
            If Round(v, LEN_ROUND) >= 0 Then
                drel_temp("校验结果") = "正确"
            End If
        End If

        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
                glo_rel_rule_num = glo_rel_rule_num + 1
            End If
        End If
        Set drel_temp = Nothing
        
    Next
End Sub


'
' 全部一级交易对手按余额合计从大到小降序排列
'
Private Sub S0400_Rule_011()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long, j As Integer, k As Integer
    Dim v2 As Double, is_50 As Boolean
    Set d_temp = cur_drule
    

    v2 = 0
    Set drel_temp = g_rule("正确")
    drel_temp("触发指标") = "4a0000_系统重要性银行交易对手合计_余额合计"
    
    
    v = g_v(start_row, "AA")
    For i = start_row To end_row - 2 Step 11
        If i <> start_row Then
            v2 = g_v(i, "AA")
            If Round(v - v2, LEN_ROUND) < 0 Then
                drel_temp("触发指标") = g_is(i - 11, 1) & " < " & g_is(i, 1)
                drel_temp("计算值") = Round(v - v2, LEN_ROUND)
                drel_temp("校验结果") = "错误"
                Exit For
            Else
                v = v2
            End If
        End If
    Next
    

    If err.Number <> 0 Then
        drel_temp("是否执行异常") = "是"
        drel_temp("异常说明") = err.Description
        err.Clear
    End If
    If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
        If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
            Set glo_rel_rule(glo_rel_rule_num) = drel_temp
        End If
    End If
    glo_rel_rule_num = glo_rel_rule_num + 1
    Set drel_temp = Nothing
End Sub

Private Sub S0400_Rule_012()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long, j As Integer, k As Integer
    Dim v2 As Double, is_50 As Boolean
    Set d_temp = cur_drule
    
    v2 = 0

    For i = start_row To end_row - 2 Step 11
        Set drel_temp = g_rule("正确")
        drel_temp("触发指标") = g_is(i, 1) & "_余额合计"
        For j = 1 To 10
            If j = 1 Then
                v = g_v(i + j, "AA")
            Else
                v2 = g_v(i + j, "AA")
                If Round(v - v2, LEN_ROUND) < 0 Then
                    drel_temp("触发指标") = g_is(i + j - 1, 1) & " < " & g_is(i + j, 1)
                    drel_temp("计算值") = Round(v - v2, LEN_ROUND)
                    drel_temp("校验结果") = "错误"
                    Exit For
                Else
                    v = v2
                End If
            End If
        Next
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub


Private Sub S0600_Rule_001()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row - 23
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1) & "_" & CStr(drel_temp("触发指标"))

        ' G=H+I+J+K+L+M+N+O+P+R+S
        v = g_v(i, "G") - (g_v(i, "H") + g_v(i, "I") + g_v(i, "J") + g_v(i, "K") + g_v(i, "L") + g_v(i, "M") + g_v(i, "N") + g_v(i, "O") + g_v(i, "P") + g_v(i, "Q") + g_v(i, "R") + g_v(i, "S"))
        drel_temp("计算值") = v
        If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0600_Rule_002()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row - 23
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1) & "_" & CStr(drel_temp("触发指标"))

        v = g_v(i, "T") - (g_v(i, "U") + g_v(i, "V"))
        drel_temp("计算值") = v
        If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub S0600_Rule_003()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    For i = start_row To end_row - 23
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(i, 1) & "_" & CStr(drel_temp("触发指标"))

        v = g_v(i, "W") - (g_v(i, "G") + g_v(i, "T"))
        drel_temp("计算值") = v
        If Abs(Round(v, LEN_ROUND)) <= glo_thd Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub


Private Sub S0600_Rule_004()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    Dim s61 As Long, s62 As Long
    s62 = 0
    s61 = 0
    For i = end_row To start_row Step -1
        If g_is(i, 1) = "620000" Then
            s62 = i
        End If
        
        If g_is(i, 1) = "610000" Then
            s61 = i
            Exit For
        End If
    Next
    
    If s61 = 0 Or s62 = 0 Or s61 > s62 Or s61 < start_row Or s62 < start_row Then
        Exit Sub
    End If
    
    v2 = 0
    Set drel_temp = g_rule("正确")
    
    v = g_v(start_row, "W")
    For i = s61 + 1 To s62 - 1 Step 1
        If i <> start_row Then
            v2 = g_v(i, "W")
            If Round(v - v2, LEN_ROUND) < 0 Then
                drel_temp("计算值") = Round(v - v2, LEN_ROUND)
                drel_temp("校验结果") = "错误"
                Exit For
            Else
                v = v2
            End If
        End If
    Next

    If err.Number <> 0 Then
        drel_temp("是否执行异常") = "是"
        drel_temp("异常说明") = err.Description
        err.Clear
    End If
    If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
        If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
            Set glo_rel_rule(glo_rel_rule_num) = drel_temp
        End If
    End If
    glo_rel_rule_num = glo_rel_rule_num + 1
    Set drel_temp = Nothing
End Sub

Private Sub S0600_Rule_005()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long
    Set d_temp = cur_drule
    
    Dim s62 As Long
    s62 = 0
    For i = end_row To start_row Step -1
        If g_is(i, 1) = "620000" Then
            s62 = i
            Exit For
        End If
    Next
    
    If s62 = 0 Or s62 < start_row Then
        Exit Sub
    End If
    
    v2 = 0
    Set drel_temp = g_rule("正确")
    
    v = g_v(start_row, "W")
    For i = s62 + 1 To end_row Step 1
        If i <> start_row Then
            v2 = g_v(i, "W")
            If Round(v - v2, LEN_ROUND) < 0 Then
                drel_temp("计算值") = Round(v - v2, LEN_ROUND)
                drel_temp("校验结果") = "错误"
                Exit For
            Else
                v = v2
            End If
        End If
    Next

    If err.Number <> 0 Then
        drel_temp("是否执行异常") = "是"
        drel_temp("异常说明") = err.Description
        err.Clear
    End If
    If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
        If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
            Set glo_rel_rule(glo_rel_rule_num) = drel_temp
        End If
    End If
    glo_rel_rule_num = glo_rel_rule_num + 1
    Set drel_temp = Nothing
End Sub


Private Sub S0500_Rule_001()
    On Error Resume Next
    
    Dim v As Double, drel_temp As Object, d_temp As Object, i As Long, j As Integer
    Set d_temp = cur_drule
    
    For j = 4 To end_col
        Set drel_temp = g_rule()
        drel_temp("触发指标") = g_is(CLng(col_row), j) & "_" & g_is(CLng(col_row) + 2, j) & "_" & g_is(CLng(col_row) + 3, j) & "_" & CStr(drel_temp("触发指标"))
        v = 0
        
        For i = start_row To end_row
            v = v + g_iv(i, j)
        Next
        
        drel_temp("计算值") = v
        If (Round(v, LEN_ROUND)) >= 0 And (Round(v, LEN_ROUND)) <= 100 Then
            drel_temp("校验结果") = "正确"
        Else
            
        End If
        
        If err.Number <> 0 Then
            drel_temp("是否执行异常") = "是"
            drel_temp("异常说明") = err.Description
            err.Clear
        End If
        
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
            End If
        End If
        glo_rel_rule_num = glo_rel_rule_num + 1
        Set drel_temp = Nothing
    Next
End Sub

Private Sub Add_Right_Rule_Record(Optional jybd As String = "")
    Dim drel_temp As Object
    If glo_isRight = True Then
        Set drel_temp = g_rule(正确)
        If jybd <> "" Then drel_temp("校验表单") = jybd
        If (is_only_err = True And CStr(drel_temp("校验结果")) = "错误") Or is_only_err = False Then
            If Not (CStr(d_temp("是否仅显示错误")) <> "" And CStr(drel_temp("校验结果")) = "正确") Then
                Set glo_rel_rule(glo_rel_rule_num) = drel_temp
                glo_rel_rule_num = glo_rel_rule_num + 1
            End If
        End If
    End If
End Sub
    
