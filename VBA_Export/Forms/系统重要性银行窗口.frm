VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 系统重要性银行窗口 
   Caption         =   "系统重要性银行审核"
   ClientHeight    =   4455
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   8052
   OleObjectBlob   =   "系统重要性银行窗口.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "系统重要性银行窗口"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private arrCmd(1 To 1) As CMySelCmdb


Private Sub CheckBox1_Click()
    Call gloCft.SetConfig("系统重要性-" & Me.CheckBox1.name & "-历史选项", CStr(Me.CheckBox1.value))
End Sub

Private Sub CheckBox2_Click()
    Call gloCft.SetConfig("系统重要性-" & Me.CheckBox2.name & "-历史选项", CStr(Me.CheckBox2.value))
End Sub

Private Sub CheckBox3_Click()
    Call gloCft.SetConfig("系统重要性-" & Me.CheckBox3.name & "-历史选项", CStr(Me.CheckBox3.value))
End Sub

Private Sub cmd_sp_Click()
    Dim savePath As String
    savePath = gloFt.GetOutputFile
    Me.tb_sp.value = savePath
    
    Call gloCft.SetConfig("系统重要性-" & Me.tb_sp.name & "-历史选项", CStr(Me.tb_sp.value))
End Sub

'
'   清空文件选择框
'
Private Sub CommandButton1_Click()
    Me.tb1.value = ""
    Me.tb_sp.value = ""
    CheckBox1.value = True
    CheckBox2.value = True
    CheckBox3.value = True
    
    Call SaveSelectFiles
End Sub

Private Sub SaveSelectFiles()
    
    Call gloCft.SetConfig("系统重要性-" & Me.tb1.name & "-历史选项", CStr(Me.tb1.value))
    Call gloCft.SetConfig("系统重要性-" & Me.tb_sp.name & "-历史选项", CStr(Me.tb_sp.value))
    Call gloCft.SetConfig("系统重要性-" & Me.CheckBox1.name & "-历史选项", CStr(Me.CheckBox1.value))
    Call gloCft.SetConfig("系统重要性-" & Me.CheckBox2.name & "-历史选项", CStr(Me.CheckBox2.value))
    Call gloCft.SetConfig("系统重要性-" & Me.CheckBox3.name & "-历史选项", CStr(Me.CheckBox3.value))
End Sub

Private Sub Userform_Initialize()
    On Error Resume Next
    Dim units As Variant, u As Variant
    Dim cmdt As CMySelCmdb
    
    Dim i As Integer
    For i = 1 To 1 Step 1
        Set cmdt = New CMySelCmdb
        Set cmdt.cmd = Me.Controls("cmd" & CStr(i))
        Set cmdt.frmT = Me
        Set arrCmd(i) = cmdt
    Next i
    
    Me.tb1.value = gloCft.GetConfig("系统重要性-" & Me.tb1.name & "-历史选项")
    Me.tb_sp.value = gloCft.GetConfig("系统重要性-" & Me.tb_sp.name & "-历史选项")
    CheckBox1.value = gloCft.GetConfig("系统重要性-" & Me.CheckBox1.name & "-历史选项")
    CheckBox2.value = gloCft.GetConfig("系统重要性-" & Me.CheckBox2.name & "-历史选项")
    CheckBox3.value = gloCft.GetConfig("系统重要性-" & Me.CheckBox3.name & "-历史选项")
    
    Set cmdt = Nothing
    
    耗时标签.Visible = False

End Sub

Private Sub 执行核查_Click()
    耗时标签.Visible = True
    
    
    
    Dim savePath As String, r As Variant, tip As String
    savePath = Me.tb_sp.value
    Call SaveSelectFiles
    If Me.tb1.value = "" Or savePath = "" Then
        MsgBox "本期文件或输出文件漏选"
    ElseIf CheckBox1.value = False And CheckBox2.value = False And CheckBox3.value = False Then
        MsgBox "未选择任一数据批次"
    Else
        tip = ""
        glo_data_pc = ""
        If CheckBox1.value = True Then
            glo_data_pc = glo_data_pc & "1批"
            tip = "1批"
        End If
        If CheckBox2.value = True Then
            glo_data_pc = glo_data_pc & "2批"
            If tip = "" Then
                tip = "2批"
            Else
                tip = tip & "," & "2批"
            End If
        End If
        If CheckBox3.value = True Then
            glo_data_pc = glo_data_pc & "3批"
            If tip = "" Then
                tip = "3批"
            Else
                tip = tip & "," & "3批"
            End If
        End If
        
        If IsContinue("是否执行，核查数据批次范围包括（" & tip & ")") = True Then
            Call 系统重要行银行审核
            耗时标签.Visible = False
            Unload Me
        End If
    End If
End Sub
