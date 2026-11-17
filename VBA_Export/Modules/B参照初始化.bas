Attribute VB_Name = "B参照初始化"
Public Sub Init_Cz()
    Call Init_Cz_1
End Sub

'
'   初始化报表参照、指标参照
'
Public Sub Init_Cz_1()
    Set glo_cz_table = CreateObject("scripting.dictionary")
    Set glo_cz_target = CreateObject("scripting.dictionary")
    Set glo_cz_target_row = CreateObject("scripting.dictionary")
    Set glo_cz_target_col = CreateObject("scripting.dictionary")
    Set dRuleRef = CreateObject("scripting.dictionary")
    
    Dim ws As Worksheet, arr As Variant, i As Long, j As Integer, d As Object
    Dim d_t As Object
    Set ws = ThisWorkbook.Worksheets("报表参照")
    arr = gloEt.GetUsedRange(ws).value
    
    
    For i = 2 To UBound(arr)
        Set d = CreateObject("scripting.dictionary")
        For j = 1 To UBound(arr, 2)
            d(CStr(arr(1, j))) = CStr(arr(i, j))
        Next
        Set glo_cz_table(CStr(arr(i, 1))) = d
        Set d = Nothing
    Next
    
    Set ws = ThisWorkbook.Worksheets("指标参照")
    arr = gloEt.GetUsedRange(ws).value
    
    For i = 2 To UBound(arr)
        If InStr(CStr(arr(i, 9)), "否") <= 0 Then
            glo_cz_target_row(CStr(arr(i, 1)) & "_" & YcTarget(CStr(arr(i, 3)))) = 1
            glo_cz_target_col(CStr(arr(i, 1)) & "_" & CStr(arr(i, 5))) = 1
            Set d = CreateObject("scripting.dictionary")
            For j = 1 To UBound(arr, 2)
                d(CStr(arr(1, j))) = CStr(arr(i, j))
            Next
            Set glo_cz_target(CStr(arr(i, 7))) = d
            Set d = Nothing
        End If
    Next
    
    Set ws = ThisWorkbook.Worksheets("函数校验")
    arr = gloEt.GetUsedRange(ws).value
    For i = 2 To UBound(arr)
        If InStr(CStr(arr(i, 3)), "否") <= 0 Then
            If CStr(arr(i, 5)) <> "" Then
                If InStr(glo_data_pc, CStr(arr(i, 4))) > 0 Then
                    If dRuleRef.exists(CStr(arr(i, 1))) Then
                        Set d = dRuleRef(CStr(arr(i, 1)))
                    Else
                        Set d = CreateObject("scripting.dictionary")
                    End If
                    Set d_t = CreateObject("scripting.dictionary")
                    For j = 1 To UBound(arr, 2)
                        d_t(CStr(arr(1, j))) = CStr(arr(i, j))
                    Next
                    ' 每个报表、每个函数名唯一
                    Set d(CStr(arr(i, 5))) = d_t
                    Set dRuleRef(CStr(arr(i, 1))) = d
                    Set d = Nothing
                    Set d_t = Nothing
                End If
            End If
        End If
    Next
End Sub


'
'   获取参照信息
'
Public Function Find_Cz(d As Object, key As String, v_key As String) As Variant
    Dim v As Variant, d_t As Object
    v = ""
    If d.exists(key) Then
        Set d_t = d(key)
        If d_t.exists(v_key) Then
            v = d_t(v_key)
        Else
            v = ""
        End If
    Else
        v = ""
    End If
    Find_Cz = v
End Function

'
'   指标代码生成规则，表格_行（少于3位的补齐）
'
Public Function YcGenTarget(table As String, row_code As String, col_code As String) As String
    On Error Resume Next
    If Len(row_code) < 3 Then
        row_code = Format(row_code, "000")
    End If
    YcGenTarget = table & "_" & row_code & "_" & col_code
End Function

Public Function YcTarget(target_code As String) As String
    On Error Resume Next
    If Len(target_code) < 3 Then
        target_code = Format(target_code, "000")
    End If
    YcTarget = target_code
End Function
