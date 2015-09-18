Attribute VB_Name = "SELECT_File"
'------------------------
'Purpose:Select A XML File To Create One XML Dom Object
'Version:
'2015-6-26  re-design version
'------------------------
Public xml_filepath As String
Public xml_filename As String
Public xmlDom As MSXML2.DOMDocument
Public mapping_select_file_flg As Integer
Public src_select_file_flg As Integer
Public tgt_select_file_flg As Integer
Public cancel_selection As Integer

Public Sub Sub_Select_File()
'On Error GoTo FATAL_ERROR
    'Only keep one file in memory
    mapping_select_file_flg = 0
    src_select_file_flg = 0
    tgt_select_file_flg = 0
    
    'Create browser to select a file
    Set objFl = Application.FileDialog(msoFileDialogFilePicker)

    With objFl
        If .Show = -1 Then
            xml_filepath = .SelectedItems(1)
            cancel_selection = 0
        Else
        'Click cancel
            cancel_selection = 1
            Exit Sub
        End If
    End With
    If xml_filepath <> "" Then
            Select Case ConsoleForm.Console_MultiPage.Value
                Case "0"
                    mapping_select_file_flg = 1
                Case "1"
                    src_select_file_flg = 1
                Case "2"
                    tgt_select_file_flg = 1
            End Select
        Set FSO = CreateObject("Scripting.FileSystemObject")
        xml_filename = FSO.GetBaseName(xml_filepath) + ".xml"
        xml_filepath = FSO.GetParentFolderName(xml_filepath)
        Set FSO = Nothing
    End If
    Set objFl = Nothing
    
    'If dtd exist in the path?
    If Dir(xml_filepath + "/powrmart.dtd") = "" Then
        Call Sub_FatalError_Msgbox("Please put powrmart.dtd under you working directory.")
        Exit Sub
    End If
    
    'Loda xml file
    Set xmlDom = New MSXML2.DOMDocument
    MsgBox xml_filepath + "/" + xml_filename
    If Not xmlDom.Load(xml_filepath + "/" + xml_filename) Then
         Call Sub_FatalError_Msgbox("XML file has syntax error." + vbLf + "MLUO: This might cause by Informatica export wizzard bug. Recommend to utilize plugin named 'XML Tools' of Notepad++ to check XML syntax.")
         Exit Sub
    End If
    
    Select Case ConsoleForm.Console_MultiPage.Value
    Case "0"
        Call Sub_Hint_Box_Set(Format(Time, "hh:mm:ss") + ": You have loaded the xml file of mapping successfully and all transformations have been listed at right." + vbLf)
        Call Sub_Hint_Box_Add("1. To edit a specified transformation, select a transformation name at the right of present worksheet, then click 'Edit This Transformation'." + vbLf)
        Call Sub_Hint_Box_Add("2. To edit links between two transformations, select two transformation names at the right of present worksheet, then click 'Goto AutoLink'." + vbLf)
        Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
    Case "1"
        Call Sub_Hint_Box_Set(Format(Time, "hh:mm:ss") + ": You have loaded the xml file of source successfully and the layout displayed at present worksheet." + vbLf)
        Call Sub_Hint_Box_Add("1. You can edit the layout or copy a layout from 'Layout Hygiene' tab. After modification complete, click 'Update This Source' to save change." + vbLf)
        Call Sub_Hint_Box_Add("2. To propagate some ports to other transformation, select some in worksheet, then click 'Keep These Ports For Propagation'." + vbLf)
        Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
    Case "2"
        Call Sub_Hint_Box_Set(Format(Time, "hh:mm:ss") + ": You have loaded the xml file of target successfully and the layout displayed at present worksheet." + vbLf)
        Call Sub_Hint_Box_Add("You can edit the layout or copy a layout from 'Layout Hygiene' tab. After modification complete, click 'Update This Target' to save change." + vbLf)
        Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
    End Select
    
    Exit Sub
'FATAL_ERROR:
'    Call Sub_Error_Handle("Sub_Select_File")
End Sub
