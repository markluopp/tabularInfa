Private Sub worksheet_change(ByVal Target As Range)
    If Target.Address(0, 0) = "A2" Or Target.Address(0, 0) = "B4" Then
        select_file_flg = 0
    End If
End Sub
Private Sub worksheet_activate()
    ConsoleForm.Console_MultiPage.Value = 1
End Sub