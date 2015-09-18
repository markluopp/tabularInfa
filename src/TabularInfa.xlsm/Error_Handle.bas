Attribute VB_Name = "Error_Handle"
'----------------------------------
'Purpose:Handle Fatal Error When Excute A Function
'Version:
'2015-6-26 Initial Version
'----------------------------------
Public Sub Sub_Error_Handle(sub_name As String)
    'Reset all public objects
    Set xmlDom = Nothing
    ReDim src_keep_port_name(0)
    ReDim src_keep_port_data_type(0)
    ReDim src_keep_port_prec(0)
    ReDim src_keep_port_scale(0)
    
    MsgBox "Fatal Error Occurs When Excute " + sub_name + ", Please Contact To mluo@merkleinc.com.", vbExclamation, "Fatal Error"
    Exit Sub
End Sub

Public Sub Sub_Hint_Box_Add(msg As String)
    ConsoleForm.HintTextBox.Text = ConsoleForm.HintTextBox.Text + msg
End Sub

Public Sub Sub_Hint_Box_Set(msg As String)
    ConsoleForm.HintTextBox.Text = msg
End Sub

Public Sub Sub_OkOnly_Msgbox(msg As String)
    MsgBox msg, vbOKOnly, "TabularInfa"
End Sub

Public Sub Sub_FatalError_Msgbox(msg As String)
    MsgBox msg, vbExclamation, "Fatal Error"
End Sub

'Sub test()
'    Call Sub_Error_Handle("test")
'End Sub
