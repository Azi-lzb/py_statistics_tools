Attribute VB_Name = "A系统重要性银行审核"
Public arrCur As Variant    ' 当前所有数据
Public glo_wb As Workbook   ' 目标工作簿
Public glo_ws As Worksheet   ' 目标工作表
Public glo_aData As Object  '当前文件所有报表数据
Public glo_cz_target As Object
Public glo_cz_target_row As Object
Public glo_cz_target_col As Object
Public glo_cz_table As Object
Public glo_d_data As Object
Public glo_data_pc As String    '数据批次
Public glo_org As Object


Public dSpcRef As Object  '特殊指标表  字典

Public dRuleRef As Object
Public dComplexRef As Object   '复杂指标表 字典
Private dComplexChecked As Object   '用于存储已校验过的规则

Public dCurDataRef As Object  '当期表  字典
Public dPreDataRef As Object  '上期表  字典

Public glo_rel_rule As Object
Public glo_rel_rule_num As Long
Public glo_thd As Double
Public glo_isRight As Boolean

Public Const LEN_ROUND As Integer = 6

Public Sub EraseSomthing()
    Set glo_d_data = Nothing
    Set glo_wb = Nothing
    Set glo_ws = Nothing
    Set glo_cz_target = Nothing
    Set glo_cz_target_row = Nothing
    Set glo_cz_target_col = Nothing
    Set glo_cz_table = Nothing
    
    Set glo_org = Nothing
    Set dComplexRef = Nothing
    Set dCurDataRef = Nothing
    Set dPreDataRef = Nothing
    Set glo_aData = Nothing
    Set glo_rel_rule = Nothing
    glo_rel_rule_num = 0
End Sub

Public Sub show_xtzyx()
    Call 系统重要性银行窗口.Show(0)
End Sub
    


Public Sub 系统重要行银行审核()
    Call EraseSomthing
    Call Init_Cz
    Call InitComplexRef
    glo_rel_rule_num = 1
    Set glo_rel_rule = CreateObject("scripting.dictionary")
    Dim fn As String, fns As Variant, y As Integer
    fn = CStr(系统重要性银行窗口.tb1.value)
    gloT = Timer
    If fn <> "" Then
        fns = Split(fn, ",")
        For y = 0 To UBound(fns)
            fn = CStr(fns(y))
            'fn = "C:\Users\caoya\Desktop\2026-07-13 系统性重要性银行\reports#911000001000013428#2025-12-31#01#中国银行股份有限公司（缺表5）.xls"
        
            Set glo_wb = gloEt.GetWb(fn)
            Set glo_aData = CreateObject("scripting.dictionary") ' 单家机构涉及的所有表
            Call Init_Data  ' 读取了单个文件所有数据
            Call CheckData_RuleCz   ' 执行配置中的
            Set glo_aData = Nothing
        Next
        
    End If
    Call CheckData_ComplexRef
    
    Call Output_ChecktData_ToWs
    
    ThisWorkbook.Worksheets("校验结果").Activate
    MsgBox "执行结束。共耗时:" & CStr(Timer - gloT) & "秒。"
End Sub

'
'   将字典校验结果数据转换为数组
'
Public Sub Output_ChecktData_ToWs()
    Dim dDataCol As Object, i As Long, v As Variant, k As Variant, arr As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("校验结果")
    ws.Cells.ClearContents
    
    
    
    Call gloCft.ShowStatusMessage_Up("公式校验完毕，转换为数组。")
    ReDim arr(1 To 1 + glo_rel_rule.count, 1 To 24) As Variant
    i = 0
    i = i + 1: arr(1, i) = "校验表单"
    i = i + 1: arr(1, i) = "机构代码"
    i = i + 1: arr(1, i) = "机构名称"
    i = i + 1: arr(1, i) = "触发指标"
    i = i + 1: arr(1, i) = "规则类型"
    i = i + 1: arr(1, i) = "校验描述"
    i = i + 1: arr(1, i) = "校验结果"
    i = i + 1: arr(1, i) = "计算值"
    i = i + 1: arr(1, i) = "是否执行异常"
    i = i + 1: arr(1, i) = "异常说明"
    i = i + 1: arr(1, i) = "校验规则"
    i = i + 1: arr(1, i) = "校验过程"
    i = i + 1: arr(1, i) = "规则函数名"
    'i = i + 1: arr(1, i) = "计算值过程"
    
    Set dDataCol = CreateObject("scripting.dictionary")
    For i = 1 To UBound(arr, 2)
        dDataCol(CStr(arr(1, i))) = i
    Next
    
    If glo_rel_rule.count > 0 Then
        
        i = 2 '
        For Each v In glo_rel_rule.items
            If v Is Nothing Then
            
            Else
                For Each k In v.keys
                    If dDataCol.exists(k) Then
                        arr(i, dDataCol(k)) = v(k)  ' 填充对应指标的值
                    End If
                Next
                i = i + 1
            End If
            
            
            
            If i Mod 5000 = 0 Then Call gloCft.ShowStatusMessage_Up("正在进行数组转换，执行到第" & CStr(i) & "行。")
        Next
    End If
    
    
    ws.Range("A1").Resize(UBound(arr), UBound(arr, 2)).NumberFormat = "@"
    ws.Range("A1").Resize(UBound(arr), UBound(arr, 2)).value = arr
    Call gloCft.ShowStatusMessage_Up("")
End Sub






Sub Task_B()
    MsgBox "执行B"
End Sub

Sub Task_C()
    MsgBox "执行C"
End Sub

Sub RunTasks()

    Dim task As Variant, arrs As Variant
    arrs = Split("A,B,C", ",")
    
    For Each task In arrs
        
        Application.Run "Task_" & task
    Next

End Sub

Function IsEnabled(taskName As String) As Boolean

    Dim ws As Worksheet
    Set ws = Sheets("配置")
    
    IsEnabled = ws.Range("B" & _
        Application.Match(taskName, ws.Columns("A"), 0)).value = 1

End Function
