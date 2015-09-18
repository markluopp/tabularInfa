VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConsoleForm 
   Caption         =   "TabularInfa Console"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   OleObjectBlob   =   "ConsoleForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "ConsoleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------
'Purpose:
'Version:
'2015-6-26 Initial Version
'----------------------------------

'----------------------------------
'Minimize the form
'----------------------------------
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
'Private Const WS_MAXIMIZEBOX As Long = &H10000



'----------------------------------
'Switch worksheets with MultiPage change
'----------------------------------
Private Sub Console_MultiPage_Change()
    Select Case Console_MultiPage.Value
    Case 0
        ThisWorkbook.Sheets("edit_mapping").Activate
        ConsoleForm.HintTextBox.Text = Format(Time, "hh:mm:ss") + ": You are now in edit_mapping tab which is used to edit a XML file of a mapping."
        If xmlDom Is Nothing Or mapping_select_file_flg = 0 Then
            ConsoleForm.HintTextBox.Text = ConsoleForm.HintTextBox.Text + "Please Click 'Select A File' first to choose a XML file." + vbLf
        Else
            ConsoleForm.HintTextBox.Text = ConsoleForm.HintTextBox.Text + xml_filename + " has been loaded." + vbLf
        End If
        ConsoleForm.HintTextBox.Text = ConsoleForm.HintTextBox.Text + "------------------------------------------------------" + vbLf
    Case 1
        ThisWorkbook.Sheets("edit_src").Activate
        ConsoleForm.HintTextBox.Text = Format(Time, "hh:mm:ss") + ": You are now in edit_src tab which is used to edit a XML file of a source."
        If xmlDom Is Nothing Or src_select_file_flg = 0 Then
            ConsoleForm.HintTextBox.Text = ConsoleForm.HintTextBox.Text + "Please Click 'Select A File' first to choose a XML file." + vbLf
        Else
            ConsoleForm.HintTextBox.Text = ConsoleForm.HintTextBox.Text + xml_filename + " has been loaded." + vbLf
        End If
        ConsoleForm.HintTextBox.Text = ConsoleForm.HintTextBox.Text + "------------------------------------------------------" + vbLf
    Case 2
        ThisWorkbook.Sheets("edit_tgt").Activate
        ConsoleForm.HintTextBox.Text = Format(Time, "hh:mm:ss") + ": You are now in edit_tgt tab which is used to edit a XML file of a target."
        ConsoleForm.HintTextBox.Text = Format(Time, "hh:mm:ss") + ": You are now in edit_src tab which is used to edit a XML file of a source."
        If xmlDom Is Nothing Or tgt_select_file_flg = 0 Then
            ConsoleForm.HintTextBox.Text = ConsoleForm.HintTextBox.Text + "Please Click 'Select A File' first to choose a XML file." + vbLf
        Else
            ConsoleForm.HintTextBox.Text = ConsoleForm.HintTextBox.Text + xml_filename + " has been loaded." + vbLf
        End If
        ConsoleForm.HintTextBox.Text = ConsoleForm.HintTextBox.Text + "------------------------------------------------------" + vbLf
    Case 3
        ThisWorkbook.Sheets("autolink").Activate
        If mapping_select_file_flg = 0 Or xmlDom Is Nothing Then
            Sub_OkOnly_Msgbox ("Please Select Two Transformations In 'edit_mapping' Tab First.")
            ConsoleForm.HintTextBox.Text = Format(Time, "hh:mm:ss") + ": You are now in autolink tab. Please click 'Select A File' in 'edit_mapping' first!" + vbLf
        Else
            ConsoleForm.HintTextBox.Text = Format(Time, "hh:mm:ss") + ": You are now in autolink tab." + xml_filename + " is loaded. Please Select Two Transformations In 'edit_mapping' Tab First." + vbLf

        End If
        ConsoleForm.HintTextBox.Text = ConsoleForm.HintTextBox.Text + "------------------------------------------------------" + vbLf
    Case 4
        ThisWorkbook.Sheets("Layout Hygiene").Activate
        ConsoleForm.HintTextBox.Text = Format(Time, "hh:mm:ss") + ": You are now in Layout Hygiene tab." + vbLf
        ConsoleForm.HintTextBox.Text = ConsoleForm.HintTextBox.Text + "------------------------------------------------------" + vbLf
    End Select
End Sub


'----------------------------------
'Make HintBox focus at last row
'----------------------------------
Private Sub HintTextBox_Change()
    HintTextBox.SelStart = Len(HintTextBox.Text)
    HintTextBox.SelLength = 0
    HintTextBox.SetFocus
End Sub




'----------------------------------
'Button click in Mapping page
'----------------------------------
Private Sub mapping_edit_Click()
    Call Sub_Edit_Selected_Trnsf(xmlDom)
End Sub
Private Sub mapping_goto_autolink_Click()
    Call Sub_Goto_Autolink
    'Switch to 'AutoLink' tab
    Console_MultiPage.Value = 3
    Call Sub_Edit_Link(xmlDom)
End Sub
Private Sub mapping_locate_Click()
    Call Sub_Locate_Port
End Sub
Private Sub mapping_normalizer_Click()
    Call Sub_Normal_Nrm
End Sub
Private Sub mapping_propagate_Click()
    Call Sub_Prepare_Propagate(xmlDom)
End Sub
Private Sub mapping_selectfile_Click()
    Call Sub_Select_File
    If cancel_selection = 0 Then
        Call Sub_Analyse_Mapping(xmlDom)
        Call Sub_Calculate_Buffer(xmlDom)
    End If
End Sub
Private Sub mapping_update_Click()
    Call Sub_Update_Selected_Trnsf(xmlDom)
End Sub
Private Sub factor_Change()
    If mapping_select_file_flg = 1 And Not xmlDom Is Nothing Then
        Call Sub_Calculate_Buffer(xmlDom)
    End If
End Sub
Private Sub mapping_add_footprint_Click()
    Call Sub_Add_Footprint_Ports
End Sub



'----------------------------------
'Button click in Source page
'----------------------------------
Private Sub source_keep_for_propagate_Click()
    Call Sub_Keep_For_Propagate
End Sub

Private Sub source_recover_currfilename_Click()
    Call Sub_Recover_CurrFileName
End Sub

Private Sub source_select_file_Click()
    Call Sub_Select_File
    If cancel_selection = 0 Then
        Call Sub_Edit_Src(0, "", xmlDom)
    End If
End Sub
Private Sub source_update_Click()
    Call Sub_Update_Src(xmlDom)
End Sub



'----------------------------------
'Button click in Target page
'----------------------------------
Private Sub target_select_file_Click()
    Call Sub_Select_File
    If cancel_selection = 0 Then
        Call Sub_Edit_Tgt(0, "", xmlDom)
    End If
End Sub
Private Sub target_update_Click()
    Call Sub_Update_Tgt(xmlDom)
End Sub
Private Sub target_add_footprint_Click()
    Call Sub_Add_Footprint_Ports
End Sub



'----------------------------------
'Button click in AutoLink page
'----------------------------------
Private Sub autolink_autolink_Click()
    Call Sub_Autolink
End Sub
Private Sub autolink_update_Click()
    Call Sub_Update_Autolink(xmlDom)
End Sub



'----------------------------------
'Button click in Layout Hygiene page
'----------------------------------
Private Sub LH_burst_Click()
    Call Sub_Burst_Cell
End Sub
Private Sub LH_clear_Click()
    Call Sub_Clear_History
End Sub
Private Sub LH_goto_edit_tgt_Click()
    Call Sub_Goto_Edit_Tgt
End Sub
Private Sub LH_goto_editsrc_Click()
    Call Sub_Goto_Edit_Src
End Sub
Private Sub LH_hygiene_Click()
    Call Sub_Layout_Hygiene
End Sub
Private Sub OptionButton_FLD_Input_Click()
    ThisWorkbook.Sheets("Layout Hygiene").Range("A1").Value = "FLD Style"
End Sub
Private Sub OptionButton_STT_Input_Click()
    ThisWorkbook.Sheets("Layout Hygiene").Range("A1").Value = "STT Style"
End Sub
Private Sub OptionButton_Locate_Input_Click()
    ThisWorkbook.Sheets("Layout Hygiene").Range("A1").Value = "Column List To Locate"
End Sub
Private Sub OptionButton_SrcTgt_Output_Click()
    ThisWorkbook.Sheets("Layout Hygiene").Range("H1").Value = "Source/Target Style"
End Sub
Private Sub OptionButton_Trnsf_Output_Click()
    ThisWorkbook.Sheets("Layout Hygiene").Range("H1").Value = "Transformation Style"
End Sub




'----------------------------------
'Form initialize
'----------------------------------
Private Sub UserForm_Initialize()
  Dim hWndForm As Long
  Dim IStyle As Long
  hWndForm = FindWindow("ThunderDFrame", Me.Caption)
  IStyle = GetWindowLong(hWndForm, GWL_STYLE)
  IStyle = IStyle Or WS_THICKFRAME
  IStyle = IStyle Or WS_MINIMIZEBOX
  'IStyle = IStyle Or WS_MAXIMIZEBOX
  SetWindowLong hWndForm, GWL_STYLE, IStyle
  
  ConsoleForm.Console_MultiPage.layouthygiene.OptionButton_FLD_Input.Value = True
  ConsoleForm.Console_MultiPage.layouthygiene.OptionButton_SrcTgt_Output = True
  ConsoleForm.Console_MultiPage.Mapping.factor.Text = "0.9"
End Sub

