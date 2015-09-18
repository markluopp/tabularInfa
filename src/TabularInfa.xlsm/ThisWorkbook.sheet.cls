Sub workbook_open()
    ConsoleForm.Show 0
    'Initial hint notification
    ConsoleForm.HintTextBox.Text = Format(Time, "hh:mm:ss") + ": You are now in edit_mapping tab which is used to edit a XML file of a mapping. Please click 'Select A File' first to choose a XML file." + vbLf
    ConsoleForm.HintTextBox.Text = ConsoleForm.HintTextBox.Text + "------------------------------------------------------" + vbLf
    'Initial public dynamic array
    ReDim src_keep_port_name(0)
    ReDim src_keep_port_data_type(0)
    ReDim src_keep_port_prec(0)
    ReDim src_keep_port_scale(0)
End Sub


Sub TEST()
    Dim test_dom As New MSXML2.DOMDocument
    MsgBox test_dom.Load("D:\06_Informatica\InfaTool\phaseI_TabularInfa\xml_bkp\m_mdbl_MODEL_SCORE_SAF.XML")
End Sub