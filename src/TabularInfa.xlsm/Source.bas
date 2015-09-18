Attribute VB_Name = "Source"
'----------------------------------
'mluo@merkleinc.com
'Version:
'2015-4-14 intail version
'----------------------------------
Public src_keep_port_name() As String
Public src_keep_port_data_type() As String
Public src_keep_port_prec() As String
Public src_keep_port_scale() As String
Public CurrFileName_Flg As Integer

Public Sub Sub_Keep_For_Propagate()
On Error GoTo FATAL_ERROR
    p_c = 0
    For Each rn In Selection
        If rn.Column = "1" Then
            p_c = p_c + 1
        End If
    Next
    'one more than actual port count
    ReDim src_keep_port_name(p_c)
    ReDim src_keep_port_data_type(p_c)
    ReDim src_keep_port_prec(p_c)
    ReDim src_keep_port_scale(p_c)
    
    p_c = 0
    For Each rn In Selection
        Select Case rn.Column
        Case "1"
            src_keep_port_name(p_c) = add_prefix + rn.Value
            src_keep_port_data_type(p_c) = Cells(rn.row, rn.Column + 1).Value
            src_keep_port_prec(p_c) = Cells(rn.row, rn.Column + 2).Value
            src_keep_port_scale(p_c) = Cells(rn.row, rn.Column + 3).Value
            p_c = p_c + 1
        Case "2", "3", "4"
        Case Else
            MsgBox "Please only select column_name/data_type/precision/scale!"
            Exit Sub
        End Select
    Next
    
    For i = 0 To p_c - 1
        Select Case src_keep_port_data_type(i)
        Case "datetime"
            src_keep_port_data_type(i) = "date/time"
            src_keep_port_prec(i) = 29
            src_keep_port_scale(i) = 9
        Case "number"
            src_keep_port_data_type(i) = "decimal"
        Case "int"
            src_keep_port_data_type(i) = "integer"
        Case "char", "varchar"
            src_keep_port_data_type(i) = "string"
        Case "nchar", "nvarchar"
            src_keep_port_data_type(i) = "nstring"
        End Select
    Next
    
    ThisWorkbook.Sheets("edit_mapping").Activate
    Call Sub_Hint_Box_Add(Format(Time, "hh:mm:ss") + ": You have selected " + CStr(p_c) + " ports of source to propagate. Select the transformation names, then Click 'Propagate Port'." + vbLf)
    Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
        
        Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Keep_For_Propagate")
End Sub

'----------------------------------
'Convert a source node to a range
'----------------------------------
Public Sub Sub_Edit_Src(is_mapping_flg As Integer, src_name As String, xmlDom As MSXML2.DOMDocument)
On Error GoTo FATAL_ERROR
     Dim xmlSrcNode As MSXML2.IXMLDOMNode
     Dim xmlNode As MSXML2.IXMLDOMNode
     Dim xmlSrcNodeList As MSXML2.IXMLDOMNodeList
     Dim xmlNodeList As MSXML2.IXMLDOMNodeList
     
     'Get the definition name
     If InStr(src_name, "(") <> 0 Then
        src_name = Mid(src_name, InStr(src_name, "(") + 1, Len(src_name) - InStr(src_name, "(") - 1)
     End If
            output_at_row = 10
            'Clean history
            For i = ActiveSheet.UsedRange.Rows.Count To output_at_row Step -1
                ActiveSheet.Rows(i).Delete
            Next
            
            If is_mapping_flg = 0 Then
                Set xmlSrcNode = xmlDom.selectSingleNode("//POWERMART/REPOSITORY/FOLDER/SOURCE")
                ActiveSheet.Range("B5").Value = xmlSrcNode.attributes.getNamedItem("NAME").nodeValue
                ActiveSheet.Range("B4").Value = xml_filename
            Else
                Set xmlSrcNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/SOURCE")
                For Each xmlNode In xmlSrcNodeList
                    If xmlNode.attributes.getNamedItem("NAME").nodeValue = src_name Then
                        Set xmlSrcNode = xmlNode
                    End If
                Next
                ActiveSheet.Range("B5").Value = src_name
            End If
            
            If xmlSrcNode Is Nothing Then
                Call Sub_OkOnly_Msgbox("Please select a SOURCE XML file!!")
            End If
            
            ThisWorkbook.Sheets("edit_src").Range("G7").Value = xmlSrcNode.attributes.getNamedItem("DATABASETYPE").nodeValue
            CurrFileName_Flg = 0
            Set xmlNodeList = xmlSrcNode.selectNodes("SOURCEFIELD")
            Set xmlNode = Nothing
            
            For Each xmlNode In xmlNodeList
                port_name = xmlNode.attributes.getNamedItem("NAME").nodeValue
                If port_name = "CurrentlyProcessedFileName" Then
                    CurrFileName_Flg = 1
                End If
                port_datatype = xmlNode.attributes.getNamedItem("DATATYPE").nodeValue
                port_pre = xmlNode.attributes.getNamedItem("PRECISION").nodeValue
                port_scale = xmlNode.attributes.getNamedItem("SCALE").nodeValue
                port_not_null = xmlNode.attributes.getNamedItem("NULLABLE").nodeValue
                port_key_type = xmlNode.attributes.getNamedItem("KEYTYPE").nodeValue
                port_bussiness_name = xmlNode.attributes.getNamedItem("BUSINESSNAME").nodeValue
                port_description = xmlNode.attributes.getNamedItem("DESCRIPTION").nodeValue

                ThisWorkbook.Sheets("edit_src").Range("A" & output_at_row).Value = port_name
                ThisWorkbook.Sheets("edit_src").Range("B" & output_at_row).Value = port_datatype
                ThisWorkbook.Sheets("edit_src").Range("C" & output_at_row).Value = port_pre
                ThisWorkbook.Sheets("edit_src").Range("D" & output_at_row).Value = port_scale
                ThisWorkbook.Sheets("edit_src").Range("E" & output_at_row).Value = port_not_null
                ThisWorkbook.Sheets("edit_src").Range("F" & output_at_row).Value = port_key_type
                ThisWorkbook.Sheets("edit_src").Range("G" & output_at_row).Value = port_bussiness_name
                ThisWorkbook.Sheets("edit_src").Range("H" & output_at_row).Value = port_description
                
                output_at_row = output_at_row + 1
            Next
            
        ActiveSheet.Range("B9:H" & output_at_row).Columns.AutoFit
        Set xmlNodeList = Nothing
        Set xmlNode = Nothing
        Set xmlSrcNodeList = Nothing
        Set xmlSrcNode = Nothing
        
        Call Sub_Hint_Box_Add(Format(Time, "hh:mm:ss") + ": Port layout has displayed at present worksheet.You can modify these ports as you want, then click 'Update This Source' to save changes." + vbLf)
        Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
        
        Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Edit_Src")
End Sub

'----------------------------------
'mluo@merkleinc.com
'Version:
'2015-4-14 intail version
'2015-4-28 fix bug#<SOURCEFIELD/> is last node#
'2015-4-29 check duplicated column name
'----------------------------------
Public Sub Sub_Update_Src(xmlDom As MSXML2.DOMDocument)
On Error GoTo FATAL_ERROR
     Dim xmlNode As MSXML2.IXMLDOMNode
     Dim newNode As MSXML2.IXMLDOMNode
     Dim xmlNodeList As MSXML2.IXMLDOMNodeList
     Dim chlNode As MSXML2.IXMLDOMNode
     Dim xmlSrcNode As MSXML2.IXMLDOMNode
     Dim xmlSrcNodeList As MSXML2.IXMLDOMNodeList
     
     'Check src XML DOM is vaild
     If (src_select_file_flg = 0 And selected_trnsf_type <> "Source Definition") Or tgt_select_file_flg = 1 Or xmlDom Is Nothing Then
        Call Sub_OkOnly_Msgbox("Please click 'Select A File' first.")
        Exit Sub
     End If
     
     src_name = ThisWorkbook.Sheets("edit_src").Range("B5").Value

            output_at_row = 10
            end_at_row = ThisWorkbook.Sheets("edit_src").Range("A65535").End(xlUp).row

            Set xmlSrcNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/SOURCE")
            For Each xmlNode In xmlSrcNodeList
                If xmlNode.attributes.getNamedItem("NAME").nodeValue = src_name Then
                    Set xmlSrcNode = xmlNode
                End If
            Next
            
            If xmlSrcNode Is Nothing Then
                Call Sub_OkOnly_Msgbox("Can not find the source named " + src_name)
                Exit Sub
            End If
            
            Set xmlNode = xmlSrcNode.selectSingleNode("SOURCEFIELD")
            port_offset = 0
            'MsgBox xmlNode.nodeName
            'MsgBox xmlNode.attributes.getNamedItem("FIELDNUMBER").nodeValue
            
            db_type = ThisWorkbook.Sheets("edit_src").Range("G7").Value
            If db_type <> "Flat File" Then
                Call Sub_OkOnly_Msgbox("Notice:This is a " + db_type + " source. We would't validate datatypes. Datatype validation only support Flat File source.")
            End If
            
            For output_at_row = 10 To end_at_row
                port_name = ThisWorkbook.Sheets("edit_src").Range("A" & output_at_row).Value
                'check duplicate column name
                For i = 10 To output_at_row
                    If port_name = ThisWorkbook.Sheets("edit_src").Range("A" & i).Value And i <> output_at_row Then
                        ThisWorkbook.Sheets("edit_src").Cells(i, 1).Interior.ColorIndex = 3
                        ThisWorkbook.Sheets("edit_src").Cells(output_at_row, 1).Interior.ColorIndex = 3
                        Call Sub_OkOnly_Msgbox("Duplicated column name!")
                        Exit Sub
                    End If
                Next
                port_datatype = ThisWorkbook.Sheets("edit_src").Range("B" & output_at_row).Value
                port_pre = ThisWorkbook.Sheets("edit_src").Range("C" & output_at_row).Value
                port_scale = ThisWorkbook.Sheets("edit_src").Range("D" & output_at_row).Value
                port_notnull = ThisWorkbook.Sheets("edit_src").Range("E" & output_at_row).Value
                port_keytype = ThisWorkbook.Sheets("edit_src").Range("F" & output_at_row).Value
                port_bussiness_name = ThisWorkbook.Sheets("edit_src").Range("G" & output_at_row).Value
                port_description = ThisWorkbook.Sheets("edit_src").Range("H" & output_at_row).Value


                'Check invalid datatype for Flat File source
                If db_type = "Flat File" Then
                        Select Case port_datatype
                        Case "bigint"
                            port_pre = 19
                            port_scale = 0
                        Case "datetime"
                            port_pre = 29
                            port_scale = 9
                        Case "string", "nstring", "int"
                            port_scale = 0
                        Case "double", "number"
                        Case Else
                            ThisWorkbook.Sheets("edit_src").Cells(output_at_row, "B").Interior.ColorIndex = 3
                            Call Sub_OkOnly_Msgbox("Invalid Flat File source data type '" + port_datatype + "' for informatica.")
                            Exit Sub
                        End Select
                End If
                'Check notnull
                        Select Case port_notnull
                        Case "NULL", "NOTNULL"
                        Case Else
                            ThisWorkbook.Sheets("edit_src").Cells(output_at_row, "E").Interior.ColorIndex = 3
                            Call Sub_OkOnly_Msgbox("Invalid NOT NULL value '" + CStr(port_notnull) + "' for informatica.")
                            Exit Sub
                        End Select
                'Check keytype
                        Select Case port_keytype
                        Case "NOT A KEY", "RRIMARY Key", "FOREIGN Key", "PRIMARY/FOREIGN KEY"
                        Case Else
                            ThisWorkbook.Sheets("edit_src").Cells(output_at_row, "F").Interior.ColorIndex = 3
                            Call Sub_OkOnly_Msgbox("Invalid KEY TYPE '" + CStr(port_keytype) + "' for informatica.")
                            Exit Sub
                        End Select

                        'Add port
                        If xmlNode Is Nothing Then
                            GoTo skip_node_name
                        End If
                        If xmlNode.nodeName <> "SOURCEFIELD" Then
skip_node_name:
                            Set newNode = xmlDom.createElement("SOURCEFIELD")
                                    Set src_attr = xmlDom.createAttribute("BUSINESSNAME")
                                    src_attr.Value = port_bussiness_name
                                    newNode.attributes.setNamedItem (src_attr)
                                    
                                    Set src_attr = xmlDom.createAttribute("DATATYPE")
                                    src_attr.Value = port_datatype
                                    newNode.attributes.setNamedItem (src_attr)

                                    Set src_attr = xmlDom.createAttribute("DESCRIPTION")
                                    src_attr.Value = port_description
                                    newNode.attributes.setNamedItem (src_attr)

                                    Set src_attr = xmlDom.createAttribute("FIELDNUMBER")
                                    src_attr.Value = str(output_at_row - 9)
                                    newNode.attributes.setNamedItem (src_attr)

                                    Set src_attr = xmlDom.createAttribute("FIELDPROPERTY")
                                    src_attr.Value = "0"
                                    newNode.attributes.setNamedItem (src_attr)

                                    Set src_attr = xmlDom.createAttribute("FIELDTYPE")
                                    src_attr.Value = "ELEMITEM"
                                    newNode.attributes.setNamedItem (src_attr)

                                    Set src_attr = xmlDom.createAttribute("HIDDEN")
                                    src_attr.Value = "NO"
                                    newNode.attributes.setNamedItem (src_attr)

                                    Set src_attr = xmlDom.createAttribute("KEYTYPE")
                                    src_attr.Value = port_keytype
                                    newNode.attributes.setNamedItem (src_attr)

                                    Set src_attr = xmlDom.createAttribute("LENGTH")
                                    src_attr.Value = port_pre
                                    newNode.attributes.setNamedItem (src_attr)

                                    Set src_attr = xmlDom.createAttribute("LEVEL")
                                    src_attr.Value = "0"
                                    newNode.attributes.setNamedItem (src_attr)
                                    
                                    Set src_attr = xmlDom.createAttribute("NAME")
                                    src_attr.Value = port_name
                                    newNode.attributes.setNamedItem (src_attr)
                                    
                                    Set src_attr = xmlDom.createAttribute("NULLABLE")
                                    src_attr.Value = port_notnull
                                    newNode.attributes.setNamedItem (src_attr)
                                    
                                    Set src_attr = xmlDom.createAttribute("OCCURS")
                                    src_attr.Value = "0"
                                    newNode.attributes.setNamedItem (src_attr)
                                    
                                    Set src_attr = xmlDom.createAttribute("OFFSET")
                                    src_attr.Value = port_offset
                                    newNode.attributes.setNamedItem (src_attr)
                                    
                                    Set src_attr = xmlDom.createAttribute("PHYSICALLENGTH")
                                    src_attr.Value = port_pre
                                    newNode.attributes.setNamedItem (src_attr)
                                    
                                    Set src_attr = xmlDom.createAttribute("PHYSICALOFFSET")
                                    src_attr.Value = str(port_offset)
                                    newNode.attributes.setNamedItem (src_attr)
                                    
                                    Set src_attr = xmlDom.createAttribute("PICTURETEXT")
                                    src_attr.Value = ""
                                    newNode.attributes.setNamedItem (src_attr)
                                    
                                    Set src_attr = xmlDom.createAttribute("PRECISION")
                                    src_attr.Value = port_pre
                                    newNode.attributes.setNamedItem (src_attr)

                                    Set src_attr = xmlDom.createAttribute("SCALE")
                                    src_attr.Value = port_scale
                                    newNode.attributes.setNamedItem (src_attr)
                                    
                                    Set src_attr = xmlDom.createAttribute("USAGE_FLAGS")
                                    src_attr.Value = ""
                                    newNode.attributes.setNamedItem (src_attr)
                            If xmlNode Is Nothing Then
                                Set xmlNode = xmlDom.selectSingleNode("//POWERMART/REPOSITORY/FOLDER/SOURCE")
                                xmlNode.appendChild newNode
                                Set xmlNode = xmlNode.LastChild
                            Else
                                xmlNode.parentNode.insertBefore newNode, xmlNode
                                Set xmlNode = xmlNode.previousSibling
                            End If
                        Else
                            xmlNode.attributes.getNamedItem("BUSINESSNAME").nodeValue = port_bussiness_name
                            xmlNode.attributes.getNamedItem("DESCRIPTION").nodeValue = port_description
                            xmlNode.attributes.getNamedItem("NAME").nodeValue = port_name
                            xmlNode.attributes.getNamedItem("DESCRIPTION").nodeValue = port_description
                            xmlNode.attributes.getNamedItem("DATATYPE").nodeValue = port_datatype
                            xmlNode.attributes.getNamedItem("PRECISION").nodeValue = port_pre
                            xmlNode.attributes.getNamedItem("LENGTH").nodeValue = port_pre
                            xmlNode.attributes.getNamedItem("PHYSICALLENGTH").nodeValue = port_pre
                            xmlNode.attributes.getNamedItem("SCALE").nodeValue = port_scale
                            
                            xmlNode.attributes.getNamedItem("KEYTYPE").nodeValue = port_keytype
                            xmlNode.attributes.getNamedItem("NULLABLE").nodeValue = port_notnull
                            xmlNode.attributes.getNamedItem("OFFSET").nodeValue = port_offset
                            xmlNode.attributes.getNamedItem("PHYSICALOFFSET").nodeValue = port_offset
                        End If
                            
                        port_offset = port_offset + port_pre
                        Set xmlNode = xmlNode.nextSibling

                    Next output_at_row

                    If Not xmlNode Is Nothing Then
                    'On Error GoTo sotp_delete_node
                        While xmlNode.nodeName = "SOURCEFIELD"
                            If xmlNode.nextSibling Is Nothing Then
                                xmlNode.parentNode.removeChild xmlNode
                                GoTo sotp_delete_node
                            Else
                                Set xmlNode = xmlNode.nextSibling
                                xmlNode.previousSibling.parentNode.removeChild xmlNode.previousSibling
                            End If
                        Wend
                    End If
sotp_delete_node:
            If port_name <> "CurrentlyProcessedFileName" And CurrFileName_Flg = 1 Then
                Sub_OkOnly_Msgbox ("The Source You are editting has checked 'Add Currently Processed Flat File Name Port' option. Please Click 'Recover CurrentlyProcessedFileName' First.")
                Exit Sub
            End If
            xmlDom.Save xml_filepath + "/" + xml_filename
            Set xmlNode = Nothing
            Set newNode = Nothing
            Set xmlSrcNode = Nothing
            Set xmlSrcNodeList = Nothing
            Call Sub_OkOnly_Msgbox("Complete update.")
            
            If mapping_select_file_flg = 1 And selected_trnsf_type = "Source Definition" Then
                ThisWorkbook.Sheets("edit_mapping").Activate
            End If
            
        Call Sub_Hint_Box_Add(Format(Time, "hh:mm:ss") + ": Port changes for " + src_name + " have been updated to the XML file." + vbLf)
        Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
        Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Update_Src")
End Sub

Public Sub Sub_Recover_CurrFileName()
    end_at_row = ThisWorkbook.Sheets("edit_src").Range("A65535").End(xlUp).row
    ThisWorkbook.Sheets("edit_src").Range("A" & (end_at_row + 1)).Value = "CurrentlyProcessedFileName"
    ThisWorkbook.Sheets("edit_src").Range("B" & (end_at_row + 1)).Value = "string"
    ThisWorkbook.Sheets("edit_src").Range("C" & (end_at_row + 1)).Value = "256"
    ThisWorkbook.Sheets("edit_src").Range("D" & (end_at_row + 1)).Value = "0"
    ThisWorkbook.Sheets("edit_src").Range("E" & (end_at_row + 1)).Value = "NULL"
    ThisWorkbook.Sheets("edit_src").Range("F" & (end_at_row + 1)).Value = "NOT A KEY"
    ThisWorkbook.Sheets("edit_src").Range("G" & (end_at_row + 1)).Value = ""
    ThisWorkbook.Sheets("edit_src").Range("H" & (end_at_row + 1)).Value = ""
End Sub




