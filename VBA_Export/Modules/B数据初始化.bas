Attribute VB_Name = "B数据初始化"
Public Sub Init_Data()
    
    Set glo_d_data = CreateObject("scripting.dictionary")
    Set dCurDataRef = CreateObject("scripting.dictionary")
    Set dPreDataRef = CreateObject("scripting.dictionary")
    Set glo_org = CreateObject("scripting.dictionary")
    
    'On Error Resume Next
    For Each glo_ws In glo_wb.Worksheets
        If glo_cz_table.exists(glo_ws.name) Then
            Call Init_Data_aData
            Call Init_Data_Exists_Cz
        Else
            ' 输出不存在
        End If
    Next
    
    Call Init_Data_ToArr

    
    ThisWorkbook.Worksheets("数据").Range("A1").Resize(UBound(arrCur), 6).NumberFormat = "@"
    ThisWorkbook.Worksheets("数据").Range("A1").Resize(UBound(arrCur), UBound(arrCur, 2)).value = arrCur
End Sub

'
'   先把当前工作簿所有表读取进来先
'
Private Sub Init_Data_aData()
    Dim start_row As Long, end_row As Long, col_row As Integer, t_name As String
    Dim arr As Variant, i As Long, j As Long
    Dim row_code As String, col_code As String, target_code As String
    Dim org_name As String, org_code As String, date_s As Variant, d_t As Object, t_val As Variant
    Dim data_type As String, pc As String
    
    t_name = glo_ws.name
    pc = CStr(Find_Cz(glo_cz_table, t_name, "数据批次"))
    If InStr(glo_data_pc, pc) <= 0 Then
        Exit Sub
    End If
    
    start_row = CLng(ycV(Find_Cz(glo_cz_table, t_name, "数据起始行")))
    end_row = Get_Last_Row(glo_ws)
    col_row = CLng(ycV(Find_Cz(glo_cz_table, t_name, "列编号行位置")))
    
    If start_row = 0 Or col_row = 0 Or end_row < start_row Then
        
        Exit Sub
    End If
    
    org_name = replace(CStr(glo_ws.Range(Find_Cz(glo_cz_table, t_name, "机构名称位置")).value), "填报机构名称：", "")
    org_code = replace(CStr(glo_ws.Range(Find_Cz(glo_cz_table, t_name, "社会信用代码位置")).value), "填报机构统一社会信用代码：", "")
    date_s = ycD(glo_ws.Range(Find_Cz(glo_cz_table, t_name, "数据日期位置")).value)
    glo_org(org_code) = 1
    ' 正式开始读取
    arr = gloEt.GetUsedRange(glo_ws).value
    glo_aData(t_name) = arr
End Sub
 

'
'   将字典数据转换为数组
'
Private Sub Init_Data_ToArr()
    Dim dDataCol As Object, i As Long, v As Variant, k As Variant
    
    If glo_d_data.count = 0 Then
        ReDim arrCur(-1 To -1, 1 To 10) As Variant
        Exit Sub
    End If
    
    Call gloCft.ShowStatusMessage_Up("数据读取完毕，转换为数组。")
    ReDim arrCur(1 To 1 + glo_d_data.count, 1 To 24) As Variant
    i = 0
    i = i + 1: arrCur(1, i) = "数据日期"
    i = i + 1: arrCur(1, i) = "报表代码"
    i = i + 1: arrCur(1, i) = "数值类型"
    i = i + 1: arrCur(1, i) = "机构代码"
    i = i + 1: arrCur(1, i) = "机构名称"
    i = i + 1: arrCur(1, i) = "指标代码"
    i = i + 1: arrCur(1, i) = "指标名称"
    i = i + 1: arrCur(1, i) = "当期值"
    
    Set dDataCol = CreateObject("scripting.dictionary")
    For i = 1 To UBound(arrCur, 2)
        dDataCol(CStr(arrCur(1, i))) = i
    Next
    
    
    i = 2 '
    For Each v In glo_d_data.items
        For Each k In v.keys
            arrCur(i, dDataCol(k)) = v(k)  ' 填充对应指标的值
        Next
        i = i + 1
        
        If i Mod 5000 = 0 Then Call gloCft.ShowStatusMessage_Up("正在进行数组转换，执行到第" & CStr(i) & "行。")
    Next
    
End Sub
    

'
'   参照中存在的批量读取
'
Public Sub Init_Data_Exists_Cz()
    Dim start_row As Long, end_row As Long, col_row As Integer, t_name As String
    Dim arr As Variant, i As Long, j As Long
    Dim row_code As String, col_code As String, target_code As String
    Dim org_name As String, org_code As String, date_s As Variant, d_t As Object, t_val As Variant
    Dim data_type As String, pc As String
    
    
    
    t_name = glo_ws.name
    pc = CStr(Find_Cz(glo_cz_table, t_name, "数据批次"))
    If InStr(glo_data_pc, pc) <= 0 Then
        Exit Sub
    End If
    
    start_row = CLng(ycV(Find_Cz(glo_cz_table, t_name, "数据起始行")))
    end_row = Get_Last_Row(glo_ws)
    col_row = CLng(ycV(Find_Cz(glo_cz_table, t_name, "列编号行位置")))
    
    If start_row = 0 Or col_row = 0 Or end_row < start_row Then
        
        Exit Sub
    End If
    
    org_name = replace(CStr(glo_ws.Range(Find_Cz(glo_cz_table, t_name, "机构名称位置")).value), "填报机构名称：", "")
    org_code = replace(CStr(glo_ws.Range(Find_Cz(glo_cz_table, t_name, "社会信用代码位置")).value), "填报机构统一社会信用代码：", "")
    date_s = ycD(glo_ws.Range(Find_Cz(glo_cz_table, t_name, "数据日期位置")).value)
    glo_org(org_code) = 1
    ' 正式开始读取
    If glo_aData.exists(t_name) Then
        arr = glo_aData(t_name)
    Else
        arr = gloEt.GetUsedRange(glo_ws).value
    End If
    
    
    'glo_aData(t_name) = arr
    
    For i = start_row To end_row
        row_code = t_name & "_" & YcTarget(CStr(arr(i, 1)))
        If glo_cz_target_row.exists(row_code) Then
            ' 存在才进一步读取
            For j = 1 To UBound(arr, 2)
                col_code = t_name & "_" & CStr(arr(col_row, j))
                If glo_cz_target_col.exists(col_code) Then
                    target_code = YcGenTarget(t_name, CStr(arr(i, 1)), CStr(arr(col_row, j)))
                    data_type = CStr(Find_Cz(glo_cz_target, target_code, "数值类型"))
                    t_val = arr(i, j)
                    If InStr(data_type, "文") > 0 Then
                        t_val = replace(ycTrim(t_val), "--", "")
                        If t_val = "" Then GoTo netxTarget
                    Else
                        t_val = ycV(t_val)
                        If Round(ycV(t_val), 10) = 0 Then GoTo netxTarget
                    End If
                    
                    
                    ' 增加判断指标类型,比如，百分数
                    Set d_t = CreateObject("scripting.dictionary")
                    d_t("机构名称") = org_name
                    d_t("机构代码") = org_code
                    d_t("报表代码") = t_name
                    d_t("数据日期") = date_s
                    d_t("指标代码") = target_code
                    d_t("指标名称") = CStr(Find_Cz(glo_cz_target, target_code, "指标名称"))
                    d_t("数值类型") = data_type
                    d_t("当期值") = t_val
                    
                    
                    Set glo_d_data(org_code & target_code) = d_t
                    dCurDataRef(org_code & target_code) = t_val
                    Set d_t = Nothing
                End If
netxTarget:
            Next
        End If
    Next
End Sub

'
' 获取系统性重要性银行表格最后一行
'
Public Function Get_Last_Row(ws As Worksheet) As Long
    Dim last_row As Long
    last_row = ws.Range("A" & ws.Cells.rows.count).End(xlUp).row
    
    Do While 1 > 0
        If CTR(ws.Cells(last_row, 1).value) = "" Then
            last_row = last_row - 1
        Else
            If InStr(ws.Cells(last_row, 1).value, "负责人") > 0 Or _
                InStr(ws.Cells(last_row, 1).value, "复核人") > 0 Or _
                InStr(ws.Cells(last_row, 1).value, "联系电话") > 0 Then
                last_row = last_row - 1
            ElseIf InStr(ws.Cells(last_row, 1).value, "填表人") > 0 Then
                last_row = last_row - 1
                Exit Do
            End If
        End If
    Loop
    
    Get_Last_Row = last_row
End Function

'
' 获取系统性重要性银行表格最后一列
'
Public Function Get_Last_Col(ws As Worksheet, col_row As Integer) As Long
    Dim last_col As Long
    last_col = ws.Cells(col_row, ws.Cells.Columns.count).End(xlToLeft).Column
    
    Get_Last_Col = last_col
End Function

