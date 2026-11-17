Option Explicit

Private Sub Workbook_Deactivate()
    Application.CommandBars("Worksheet menu bar").Reset
End Sub

Private Sub Workbook_Open()
    Call 修改右键菜单
End Sub

Private Sub 修改右键菜单()
    Dim ToolMenu As Object
    Dim NewMenu As Object
    With Application.CommandBars("Worksheet menu bar")
        .Reset
        '利用ID属性查找帮助菜单
        Set ToolMenu = .FindControl(id:=.Controls("工具(&T)").id)
        If ToolMenu Is Nothing Then
            '如果该菜单不存在,则将新菜单添加到末尾
            Set NewMenu = .Controls.add(Type:=msoControlPopup)
        Else
            '将新菜单添加到帮助菜单之前
            Set NewMenu = .Controls.add(Type:=msoControlPopup, _
                before:=ToolMenu.index)
        End If
        With NewMenu
            .Caption = "统计工具(&S)"
            With .Controls.add(Type:=msoControlButton)
                .Caption = "系统性重要银行核查"
                .OnAction = "show_xtzyx"
            End With
                
        End With
    End With
    Set ToolMenu = Nothing
    Set NewMenu = Nothing
End Sub

Private Sub Workbook_Activate()
    Call 修改右键菜单
End Sub
