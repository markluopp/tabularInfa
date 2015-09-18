Attribute VB_Name = "Mapping"
'------------------------
'Purpose:
'Version:
'2015-04-22  intail version
'2015-04-25  add propagate procedure
'2015-04-26  add locate procedure
'------------------------
Public selected_trnsf_name As String
'use to keep instance name of resuable transformation
Public selected_trnsf_name_hist As String
Public selected_trnsf_type As String

Public Sub Sub_Analyse_Mapping(xmlDom As MSXML2.DOMDocument)
On Error GoTo FATAL_ERROR
     Dim xmlNode As MSXML2.IXMLDOMNode
     Dim xmlNodeList As MSXML2.IXMLDOMNodeList
     
     If mapping_select_file_flg = 0 Or xmlDom Is Nothing Then
     '-----------Support Hard Code-------------
'         xml_filepath = ActiveSheet.Range("A2").Value
'         xml_filename = ActiveSheet.Range("B4").Value
'         'If dtd exist in the path?
'         If Dir(xml_filepath + "/powrmart.dtd") = "" Then
'             MsgBox "Please put powrmart.dtd under you working directory.", vbExclamation, "Fatal Error"
'             Exit Sub
'         End If
'
'         'Loda xml file
'         Set xmlDom = New MSXML2.DOMDocument
'
'         If Not xmlDom.Load(xml_filepath + "/" + xml_filename) Then
'              MsgBox "XML file has syntax error.", vbExclamation, "Fatal Error"
'              Exit Sub
'         End If
'
'         ConsoleForm.HintTextBox.Text = Format(Time, "hh:mm:ss") + ": You have loaded the xml file of mapping successfully by hard code path&name" + vbLf
'         ConsoleForm.HintTextBox.Text = ConsoleForm.HintTextBox.Text + "------------------------------------------------------" + vbLf
     '-----------------------------------------
     Else
         ActiveSheet.Range("A2").Value = xml_filepath
         ActiveSheet.Range("B4").Value = xml_filename
     End If
     
    'Clean history
     analysis_result_end_at = ActiveSheet.Range("A65535").End(xlUp).row
     If analysis_result_end_at >= 10 Then
        ActiveSheet.Range("A" + CStr(analysis_result_end_at) + ":B10").Clear
     End If
     
     Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/MAPPING/INSTANCE")
     output_at_row = 10
     For Each xmlNode In xmlNodeList
        If xmlNode.attributes.getNamedItem("REUSABLE") Is Nothing Then
            GoTo NOREUSE
        End If
        
        If xmlNode.attributes.getNamedItem("REUSABLE").nodeValue = "YES" Then
            ActiveSheet.Range("A" & output_at_row).Value = xmlNode.attributes.getNamedItem("NAME").nodeValue + "(" + xmlNode.attributes.getNamedItem("TRANSFORMATION_NAME").nodeValue + ")"
            ActiveSheet.Range("B" & output_at_row).Value = xmlNode.attributes.getNamedItem("TRANSFORMATION_TYPE").nodeValue + "(REUSABLE)"
            ActiveSheet.Range("A" & output_at_row).Font.FontStyle = "bold"
            ActiveSheet.Range("B" & output_at_row).Font.FontStyle = "bold"
        Else
NOREUSE:
            If xmlNode.attributes.getNamedItem("TRANSFORMATION_TYPE").nodeValue = "Source Definition" Or xmlNode.attributes.getNamedItem("TRANSFORMATION_TYPE").nodeValue = "Target Definition" Then
                ActiveSheet.Range("A" & output_at_row).Value = xmlNode.attributes.getNamedItem("NAME").nodeValue + "(" + xmlNode.attributes.getNamedItem("TRANSFORMATION_NAME").nodeValue + ")"
                ActiveSheet.Range("B" & output_at_row).Value = xmlNode.attributes.getNamedItem("TRANSFORMATION_TYPE").nodeValue
                ActiveSheet.Range("A" & output_at_row).Font.FontStyle = "bold"
                ActiveSheet.Range("B" & output_at_row).Font.FontStyle = "bold"
            Else
                ActiveSheet.Range("A" & output_at_row).Value = xmlNode.attributes.getNamedItem("NAME").nodeValue
                ActiveSheet.Range("B" & output_at_row).Value = xmlNode.attributes.getNamedItem("TRANSFORMATION_TYPE").nodeValue
            End If
        End If
        output_at_row = output_at_row + 1
     Next
     ActiveSheet.Range("B9:B" & output_at_row).Columns.AutoFit
     
     
     'Caculate DTM size and block size
     
     Set xmlNodeList = Nothing
     Set xmlNode = Nothing
     Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Analyse_Mapping")
End Sub


Public Sub Sub_Edit_Selected_Trnsf(xmlDom As MSXML2.DOMDocument)
On Error GoTo FATAL_ERROR
     'Check mapping XML DOM is vaild
     If mapping_select_file_flg = 0 Or xmlDom Is Nothing Then
        Call Sub_OkOnly_Msgbox("Please click 'Select A File' first.")
        Exit Sub
     End If
     
     'Disable Normalizer Button
     ConsoleForm.mapping_normalizer.Visible = False
    
    'Clean existed format
     analysis_result_end_at = ActiveSheet.Range("A65535").End(xlUp).row
     If analysis_result_end_at >= 10 Then
        ActiveSheet.Range("A" + CStr(analysis_result_end_at) + ":B10").Interior.ColorIndex = xlNone
     End If
     
     'Check selection
     If Selection.Column > 1 Or Selection.Value = "" Then
        Call Sub_OkOnly_Msgbox("Please choose a valid transformation name.")
        Exit Sub
     End If
     
     selected_trnsf_name = Selection.Value
     selected_trnsf_type = ActiveSheet.Range("B" & Selection.row).Value
     Selection.Interior.ColorIndex = 6
     
     'Highlight forward and backword transformation
     Call Sub_Highlight(selected_trnsf_name, xmlDom)
     
     header_end_at = [iv9].End(xlToLeft).Column
     selected_trnsf_name_hist = selected_trnsf_name
     'set header for different trnsf type
     Select Case selected_trnsf_type
     Case "Source Definition"
        'MsgBox "Source/Target Definition is readonly in mapping. DO NOT EDIT IT!"
        ThisWorkbook.Sheets("edit_src").Activate
         Call Sub_Hint_Box_Set(Format(Time, "hh:mm:ss") + ": Jump to 'edit_src' tab to edit the source definition " + selected_trnsf_name + " of mapping." + vbLf)
         Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
        Call Sub_Edit_Src(1, selected_trnsf_name, xmlDom)
    Case "Target Definition"
        ThisWorkbook.Sheets("edit_tgt").Activate
         Call Sub_Hint_Box_Set(Format(Time, "hh:mm:ss") + ": Jump to 'edit_tgt' tab to edit the target definition " + selected_trnsf_name + " of mapping." + vbLf)
         Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
        Call Sub_Edit_Tgt(1, selected_trnsf_name, xmlDom)
    Case "Expression", "Expression(REUSABLE)"
        If header_end_at >= 4 Then
           ActiveSheet.Range("D9:" + Chr(header_end_at + 64) + "9").Clear
        End If
        ActiveSheet.Range("D9").Value = "Port Name"
        ActiveSheet.Range("E9").Value = "Data Type"
        ActiveSheet.Range("F9").Value = "Precision"
        ActiveSheet.Range("G9").Value = "Scale"
        ActiveSheet.Range("H9").Value = "Expression"
        ActiveSheet.Range("I9").Value = "Port Type"
        ActiveSheet.Range("D9:I9").Interior.ColorIndex = 3
        ActiveSheet.Range("D9:I9").Font.FontStyle = "bold"
        'ActiveSheet.Range("D9:I9").Columns.AutoFit
        Call Sub_Edit_Exp(xmlDom, selected_trnsf_name)
        'normalizer.Enabled = True
    Case "Aggregator", "Aggregator(REUSABLE)"
        If header_end_at >= 4 Then
           ActiveSheet.Range("D9:" + Chr(header_end_at + 64) + "9").Clear
        End If
        ActiveSheet.Range("D9").Value = "Port Name"
        ActiveSheet.Range("E9").Value = "Data Type"
        ActiveSheet.Range("F9").Value = "Precision"
        ActiveSheet.Range("G9").Value = "Scale"
        ActiveSheet.Range("H9").Value = "Expression"
        ActiveSheet.Range("I9").Value = "Port Type"
        ActiveSheet.Range("J9").Value = "Group By"
        ActiveSheet.Range("D9:J9").Interior.ColorIndex = 15
        ActiveSheet.Range("D9:J9").Font.FontStyle = "bold"
        'ActiveSheet.Range("D9:I9").Columns.AutoFit
        Call Sub_Edit_Agg(xmlDom, selected_trnsf_name)
    Case "Joiner", "Joiner(REUSABLE)"
        If header_end_at >= 4 Then
           ActiveSheet.Range("D9:" + Chr(header_end_at + 64) + "9").Clear
        End If
        ActiveSheet.Range("D9").Value = "Port Name"
        ActiveSheet.Range("E9").Value = "Data Type"
        ActiveSheet.Range("F9").Value = "Precision"
        ActiveSheet.Range("G9").Value = "Scale"
        ActiveSheet.Range("H9").Value = "Port Type"
        ActiveSheet.Range("D9:H9").Interior.ColorIndex = 43
        ActiveSheet.Range("D9:H9").Font.FontStyle = "bold"
        'ActiveSheet.Range("D9:I9").Columns.AutoFit
        Call Sub_Edit_Jnr(xmlDom, selected_trnsf_name)
     Case "Normalizer", "Normalizer(REUSABLE)"
        'Enable Normalizer Button
        ConsoleForm.mapping_normalizer.Visible = True
        If header_end_at >= 4 Then
           ActiveSheet.Range("D9:" + Chr(header_end_at + 64) + "9").Clear
        End If
        ActiveSheet.Range("D9").Value = "Port Name"
        ActiveSheet.Range("E9").Value = "Data Type"
        ActiveSheet.Range("F9").Value = "Precision"
        ActiveSheet.Range("G9").Value = "Scale"
        ActiveSheet.Range("H9").Value = "Port Type"
        ActiveSheet.Range("D9:H9").Interior.ColorIndex = 12
        ActiveSheet.Range("I9").Value = "Column Name"
        ActiveSheet.Range("J9").Value = "Level"
        ActiveSheet.Range("K9").Value = "Occurs"
        ActiveSheet.Range("L9").Value = "Data Type"
        ActiveSheet.Range("M9").Value = "Prec"
        ActiveSheet.Range("N9").Value = "Scale"
        'ActiveSheet.Range("I9:N9").Interior.ColorIndex = 43
        'ActiveSheet.Range("D9:N9").Font.FontStyle = "bold"
        Call Sub_Edit_Nrm(xmlDom, selected_trnsf_name)
    Case "Sorter", "Sorter(REUSABLE)"
        If header_end_at >= 4 Then
           ActiveSheet.Range("D9:" + Chr(header_end_at + 64) + "9").Clear
        End If
        ActiveSheet.Range("D9").Value = "Port Name"
        ActiveSheet.Range("E9").Value = "Data Type"
        ActiveSheet.Range("F9").Value = "Precision"
        ActiveSheet.Range("G9").Value = "Scale"
        ActiveSheet.Range("H9").Value = "IsSortKey?"
        ActiveSheet.Range("I9").Value = "SortDirection"
        ActiveSheet.Range("D9:I9").Interior.ColorIndex = 53
        ActiveSheet.Range("D9:I9").Font.FontStyle = "bold"
        Call Sub_Edit_Srt(xmlDom, selected_trnsf_name)
     Case "Mapplet", "Mapplet(REUSABLE)"
        Call Sub_OkOnly_Msgbox("Unsupported Transformation Type.")
     Case Else
        If header_end_at >= 4 Then
           ActiveSheet.Range("D9:" + Chr(header_end_at + 64) + "9").Clear
        End If
        ActiveSheet.Range("D9").Value = "Port Name"
        ActiveSheet.Range("E9").Value = "Data Type"
        ActiveSheet.Range("F9").Value = "Precision"
        ActiveSheet.Range("G9").Value = "Scale"
        Select Case selected_trnsf_type
            Case "Source Qualifier"
                ActiveSheet.Range("D9:G9").Interior.ColorIndex = 12
            Case "Filter", "Filter(REUSABLE)"
                ActiveSheet.Range("D9:G9").Interior.ColorIndex = 9
            Case Else
                ActiveSheet.Range("D9:G9").Interior.ColorIndex = 2
        End Select
        ActiveSheet.Range("D9:G9").Font.FontStyle = "bold"
        Call Sub_Edit_Trnsf_Part(xmlDom, selected_trnsf_name)
        'MsgBox "Attention:This kind of transformation can not be updated so far."
     End Select
     Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Edit_Selected_Trnsf")
End Sub

Public Sub Sub_Update_Selected_Trnsf(xmlDom As MSXML2.DOMDocument)
On Error GoTo FATAL_ERROR
     'Check mapping XML DOM is vaild
     If mapping_select_file_flg = 0 Or xmlDom Is Nothing Then
        Call Sub_OkOnly_Msgbox("Please click 'Select A File' first.")
        Exit Sub
     End If
     
     Select Case selected_trnsf_type
     Case "Expression", "Expression(REUSABLE)"
        Call Sub_Update_Exp(xmlDom, selected_trnsf_name)
     Case "Joiner", "Joiner(REUSEABLE)"
        Call Sub_Update_Jnr(xmlDom, selected_trnsf_name)
     Case "Normalizer", "Normalizer(REUSABLE)"
        Call Sub_Update_Nrm(xmlDom, selected_trnsf_name)
     Case "Source Qualifier", "Filter"
        Call Sub_Update_Trnsf_Part(xmlDom, selected_trnsf_name)
     Case "Sorter", "Sorter(REUSABLE)"
        Call Sub_Update_Srt(xmlDom, selected_trnsf_name)
     Case "Aggregator", "Aggregator(REUSABLE)"
        Call Sub_Update_Agg(xmlDom, selected_trnsf_name)
     Case Else
        Call Sub_OkOnly_Msgbox("Unsupported Transformation Type.")
     End Select
     Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Update_Selected_Trnsf")
End Sub

Public Sub Sub_Goto_Autolink()
On Error GoTo FATAL_ERROR
    'Check mapping XML DOM is vaild
    If mapping_select_file_flg = 0 Or xmlDom Is Nothing Then
        Call Sub_OkOnly_Msgbox("Please click 'Select A File' first.")
        Exit Sub
    End If
     
    Dim rn As Range
    
    'if select two cells?
    If Selection.Count <> 2 Then
        MsgBox "Must select two transformations!"
        Exit Sub
    End If
    
    'if valid trnsf name? then assign value
    For Each rn In Selection
    
        If rn.Column <> 1 Or rn.row < 10 Or rn.Value = "" Then
            MsgBox rn.Value + "isn't a transformation name!"
            Exit Sub
        End If
        
        If fr_trnsf_name = "" Then
            fr_trnsf_name = rn.Value
        End If
        
        If fr_trnsf_name <> "" Then
            to_trnsf_name = rn.Value
        End If
    Next
    
    'switch to autolink tab
    If MsgBox("Treat " + fr_trnsf_name + " As A Start?", vbYesNo, "TabularInfa") = 6 Then
        ThisWorkbook.Sheets("autolink").Range("A2").Value = xml_filepath
        ThisWorkbook.Sheets("autolink").Range("B4").Value = xml_filename
        ThisWorkbook.Sheets("autolink").Range("B5").Value = fr_trnsf_name
        ThisWorkbook.Sheets("autolink").Range("B6").Value = to_trnsf_name
        ThisWorkbook.Sheets("autolink").Activate
        Exit Sub
    End If
    
    If MsgBox("Treat " + to_trnsf_name + " As A Start?", vbYesNo, "TabularInfa") = 6 Then
        ThisWorkbook.Sheets("autolink").Range("A2").Value = xml_filepath
        ThisWorkbook.Sheets("autolink").Range("B4").Value = xml_filename
        ThisWorkbook.Sheets("autolink").Range("B5").Value = to_trnsf_name
        ThisWorkbook.Sheets("autolink").Range("B6").Value = fr_trnsf_name
        ThisWorkbook.Sheets("autolink").Activate
        Exit Sub
    End If
    
    Call Sub_OkOnly_Msgbox("You Must Choose One Transformation As A Start!")
    Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Goto_Autolink")
End Sub


Public Sub Sub_Locate_Port()
On Error GoTo FATAL_ERROR
    header_end_at = [iv9].End(xlToLeft).Column
    port_end_at = ActiveSheet.Range("D65535").End(xlUp).row
    Dim select_cell As Range
    
    If ThisWorkbook.Sheets("Layout Hygiene").Range("A1").Value <> "Column List To Locate" Then
        Call Sub_OkOnly_Msgbox("Layout Hygiene is not ready for loacte!")
        ThisWorkbook.Sheets("Layout Hygiene").Activate
        Exit Sub
    Else
        end_at_row = ThisWorkbook.Sheets("Layout Hygiene").Range("A65535").End(xlUp).row
        For i = 3 To end_at_row
            dup_field_flg = 0
            Field = LTrim(RTrim(ThisWorkbook.Sheets("Layout Hygiene").Range("A" & i).Value))
            For j = 3 To i
                If LTrim(RTrim(ThisWorkbook.Sheets("Layout Hygiene").Range("A" & j).Value)) = Field And i <> j Then
                    dup_field_flg = 1
                End If
            Next
            If dup_field_flg = 0 Then
                match_flg = 0
                For k = 10 To port_end_at
                    If ActiveSheet.Range("D" & k).Value = Field Then
                        'select_srt = select_srt + "," + "D" + CStr(k) + ":" + Chr(header_end_at + 64) + CStr(k)
                        If select_cell Is Nothing Then
                            Set select_cell = Range("D" + CStr(k) + ":" + Chr(header_end_at + 64) + CStr(k))
                        Else
                            Set select_cell = Union(select_cell, Range("D" + CStr(k) + ":" + Chr(header_end_at + 64) + CStr(k)))
                        End If
                        match_flg = 1
                    End If
                Next
                If match_flg = 0 Then
                'can't find handle
                    If MsgBox("Can't find field named " + Field + ", Do you want continue?", vbYesNo, "TabularInfa") = vbNo Then
                         ThisWorkbook.Sheets("Layout Hygiene").Range("A" & i).Interior.ColorIndex = 3
                         ThisWorkbook.Sheets("Layout Hygiene").Activate
                         Exit Sub
                    End If
                End If
            End If
        Next
    End If
    'select_srt = Mid(select_srt, 2, Len(select_srt) - 1)
    'MsgBox select_srt
    'Range(select_srt).Select
    select_cell.Select
    Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Locate_Port")
End Sub

Public Sub Sub_Calculate_Buffer(xmlDom As MSXML2.DOMDocument)
On Error GoTo FATAL_ERROR
     'Dim xmlNode As MSXML2.IXMLDOMNode
     Dim SrcNodeList As MSXML2.IXMLDOMNodeList
     Dim SrcFieldNodeList As MSXML2.IXMLDOMNodeList
     Dim TgtNodeList As MSXML2.IXMLDOMNodeList
     Dim TgtFieldNodeList As MSXML2.IXMLDOMNodeList
     
     src_count = 0
     tgt_count = 0
     max_src_offset = 0
     max_tgt_offset = 0
     
     Set SrcNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/SOURCE")
     src_count = SrcNodeList.Length
     
     For Each SrcNode In SrcNodeList
        Set SrcFieldNodeList = SrcNode.selectNodes("SOURCEFIELD")
        src_offset = SrcFieldNodeList.Item(SrcFieldNodeList.Length - 1).attributes.getNamedItem("PHYSICALOFFSET").nodeValue
        If src_offset > max_src_offset Then
            max_src_offset = src_offset
        End If
     Next
     
     Set TgtNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/TARGET")
     tgt_count = TgtNodeList.Length
     
     For Each TgtNode In TgtNodeList
        Set TgtFieldNodeList = TgtNode.selectNodes("TARGETFIELD")
        tgt_offset = 0
        For Each TgtField In TgtFieldNodeList
            tgt_offset = tgt_offset + CInt(TgtField.attributes.getNamedItem("PRECISION").nodeValue)
        Next
        If tgt_offset > max_tgt_offset Then
            max_tgt_offset = tgt_offset
        End If
     Next
     If max_src_offset > max_tgt_offset Then
        max_offset = max_src_offset
     Else
        max_offset = max_tgt_offset
     End If
     
     block_size = max_offset * 100
     buffer_size = Round(block_size * (src_count + tgt_count) * 2 / CDbl(ConsoleForm.Console_MultiPage.Mapping.factor.Text))
     
     ConsoleForm.Console_MultiPage.Mapping.mapping_block.Text = block_size
     ConsoleForm.Console_MultiPage.Mapping.mapping_buffer.Text = buffer_size
     
    Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Locate_Port")
End Sub


