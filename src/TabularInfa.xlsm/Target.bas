Attribute VB_Name = "Target"
'----------------------------------
'mluo@merkleinc.com
'Version:
'2015-4-14 intail version
'----------------------------------

Public Sub Sub_Edit_Tgt(is_mapping_flg As Integer, tgt_name As String, xmlDom As MSXML2.DOMDocument)
On Error GoTo FATAL_ERROR
     Dim xmlNode As MSXML2.IXMLDOMNode
     Dim xmlNodeList As MSXML2.IXMLDOMNodeList
     Dim xmlTgtNode As MSXML2.IXMLDOMNode
     Dim xmlTgtNodeList As MSXML2.IXMLDOMNodeList
    
     
    'Get the definition name
     If InStr(tgt_name, "(") <> 0 Then
        tgt_name = Mid(tgt_name, InStr(tgt_name, "(") + 1, Len(tgt_name) - InStr(tgt_name, "(") - 1)
     End If
            
            output_at_row = 10
            'Clean history
            For i = ActiveSheet.UsedRange.Rows.Count To output_at_row Step -1
                ActiveSheet.Rows(i).Delete
            Next
            
            If is_mapping_flg = 0 Then
                Set xmlTgtNode = xmlDom.selectSingleNode("//POWERMART/REPOSITORY/FOLDER/TARGET")
                ActiveSheet.Range("B5").Value = xmlTgtNode.attributes.getNamedItem("NAME").nodeValue
                ActiveSheet.Range("B4").Value = xml_filename
            Else
                Set xmlTgtNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/TARGET")
                For Each xmlNode In xmlTgtNodeList
                    If xmlNode.attributes.getNamedItem("NAME").nodeValue = tgt_name Then
                        Set xmlTgtNode = xmlNode
                    End If
                Next
                ActiveSheet.Range("B5").Value = tgt_name
            End If
            
            If xmlTgtNode Is Nothing Then
                Call Sub_OkOnly_Msgbox("Please select a TARGET XML file!!")
            End If
            
            ThisWorkbook.Sheets("edit_tgt").Range("G7").Value = xmlTgtNode.attributes.getNamedItem("DATABASETYPE").nodeValue
            
            Set xmlNodeList = xmlTgtNode.selectNodes("TARGETFIELD")
            Set xmlNode = Nothing
            
            For Each xmlNode In xmlNodeList
                port_name = xmlNode.attributes.getNamedItem("NAME").nodeValue
                port_datatype = xmlNode.attributes.getNamedItem("DATATYPE").nodeValue
                port_pre = xmlNode.attributes.getNamedItem("PRECISION").nodeValue
                port_scale = xmlNode.attributes.getNamedItem("SCALE").nodeValue
                port_not_null = xmlNode.attributes.getNamedItem("NULLABLE").nodeValue
                port_key_type = xmlNode.attributes.getNamedItem("KEYTYPE").nodeValue
                port_bussiness_name = xmlNode.attributes.getNamedItem("BUSINESSNAME").nodeValue
                port_description = xmlNode.attributes.getNamedItem("DESCRIPTION").nodeValue

                            
                ThisWorkbook.Sheets("edit_tgt").Range("A" & output_at_row).Value = port_name
                ThisWorkbook.Sheets("edit_tgt").Range("B" & output_at_row).Value = port_datatype
                ThisWorkbook.Sheets("edit_tgt").Range("C" & output_at_row).Value = port_pre
                ThisWorkbook.Sheets("edit_tgt").Range("D" & output_at_row).Value = port_scale
                ThisWorkbook.Sheets("edit_tgt").Range("E" & output_at_row).Value = port_not_null
                ThisWorkbook.Sheets("edit_tgt").Range("F" & output_at_row).Value = port_key_type
                ThisWorkbook.Sheets("edit_tgt").Range("G" & output_at_row).Value = port_bussiness_name
                ThisWorkbook.Sheets("edit_tgt").Range("H" & output_at_row).Value = port_description
                
                output_at_row = output_at_row + 1
            Next
        ActiveSheet.Range("B9:H" & output_at_row).Columns.AutoFit
           Set xmlNodeList = Nothing
           Set xmlNode = Nothing
           Set xmlTgtNode = Nothing
           Set xmlTgtNodeList = Nothing
           
        Call Sub_Hint_Box_Add(Format(Time, "hh:mm:ss") + ": Port layout has displayed at present worksheet.You can modify these ports as you want, then click 'Update This Target' to save changes." + vbLf)
        Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
        Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Edit_Tgt")
End Sub

'----------------------------------
'mluo@merkleinc.com
'Version:
'2015-4-14 intail version
'2015-4-28 fix bug#<TARGETFIELD/> is last node#
'2015-4-29 check duplicated column name
'----------------------------------
Public Sub Sub_Update_Tgt(xmlDom As MSXML2.DOMDocument)
On Error GoTo FATAL_ERROR
     Dim xmlNode As MSXML2.IXMLDOMNode
     Dim newNode As MSXML2.IXMLDOMNode
     Dim xmlNodeList As MSXML2.IXMLDOMNodeList
     Dim chlNode As MSXML2.IXMLDOMNode
     Dim xmlTgtNode As MSXML2.IXMLDOMNode
     Dim xmlTgtNodeList As MSXML2.IXMLDOMNodeList
    
     'Check tgt XML DOM is vaild
     If (tgt_select_file_flg = 0 And selected_trnsf_type <> "Target Definition") Or src_select_file_flg = 1 Or xmlDom Is Nothing Then
        Call Sub_OkOnly_Msgbox("Please click 'Select A File' first.")
        Exit Sub
     End If
     
     tgt_name = ThisWorkbook.Sheets("edit_tgt").Range("B5").Value
        
            output_at_row = 10
            end_at_row = ThisWorkbook.Sheets("edit_tgt").Range("A65535").End(xlUp).row
            
            Set xmlTgtNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/TARGET")
            For Each xmlNode In xmlTgtNodeList
                If xmlNode.attributes.getNamedItem("NAME").nodeValue = tgt_name Then
                    Set xmlTgtNode = xmlNode
                End If
            Next
            
            If xmlTgtNode Is Nothing Then
                Call Sub_OkOnly_Msgbox("Can not find the target named " + src_name)
                Exit Sub
            End If

            Set xmlNode = xmlTgtNode.selectSingleNode("TARGETFIELD")
            'MsgBox xmlNode.nodeName
            'MsgBox xmlNode.attributes.getNamedItem("FIELDNUMBER").nodeValue
            
            db_type = ThisWorkbook.Sheets("edit_tgt").Range("G7").Value
            If db_type <> "Flat File" Then
                MsgBox "Notice:This is a " + db_type + " source. We would't validate datatypes. Datatype validation only support Flat File source."
            End If
            
            For output_at_row = 10 To end_at_row
                
                port_name = ThisWorkbook.Sheets("edit_tgt").Range("A" & output_at_row).Value
                 'check duplicate column name
                For i = 10 To output_at_row
                    If port_name = ThisWorkbook.Sheets("edit_tgt").Range("A" & i).Value And i <> output_at_row Then
                        ThisWorkbook.Sheets("edit_tgt").Cells(i, 1).Interior.ColorIndex = 3
                        ThisWorkbook.Sheets("edit_tgt").Cells(output_at_row, 1).Interior.ColorIndex = 3
                        Call Sub_OkOnly_Msgbox("Duplicated column name!")
                        Exit Sub
                    End If
                Next
                port_datatype = ThisWorkbook.Sheets("edit_tgt").Range("B" & output_at_row).Value
                port_pre = ThisWorkbook.Sheets("edit_tgt").Range("C" & output_at_row).Value
                port_scale = ThisWorkbook.Sheets("edit_tgt").Range("D" & output_at_row).Value
                port_notnull = ThisWorkbook.Sheets("edit_tgt").Range("E" & output_at_row).Value
                port_keytype = ThisWorkbook.Sheets("edit_tgt").Range("F" & output_at_row).Value
                port_bussiness_name = ThisWorkbook.Sheets("edit_tgt").Range("G" & output_at_row).Value
                port_description = ThisWorkbook.Sheets("edit_tgt").Range("H" & output_at_row).Value


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
                            ThisWorkbook.Sheets("edit_tgt").Cells(output_at_row, "B").Interior.ColorIndex = 3
                            Call Sub_OkOnly_Msgbox("Invalid Flat File source data type '" + port_datatype + "' for informatica.")
                            Exit Sub
                        End Select
                End If
                'Check notnull
                        Select Case port_notnull
                        Case "NULL", "NOTNULL"
                        Case Else
                            ThisWorkbook.Sheets("edit_tgt").Cells(output_at_row, "E").Interior.ColorIndex = 3
                            Call Sub_OkOnly_Msgbox("Invalid NOT NULL value '" + CStr(port_notnull) + "' for informatica.")
                            Exit Sub
                        End Select
                'Check keytype
                        Select Case port_keytype
                        Case "NOT A KEY", "RRIMARY KEY", "FOREIGN KEY", "PRIMARY/FOREIGN KEY"
                        Case Else
                            ThisWorkbook.Sheets("edit_tgt").Cells(output_at_row, "F").Interior.ColorIndex = 3
                            Call Sub_OkOnly_Msgbox("Invalid KEY TYPE '" + CStr(port_keytype) + "' for informatica.")
                            Exit Sub
                        End Select

                        'Add port
                        If xmlNode Is Nothing Then
                            Set xmlNode = xmlDom.selectSingleNode("//POWERMART/REPOSITORY/FOLDER/TARGET/TABLEATTRIBUTE")
                            If xmlNode Is Nothing Then
                                GoTo skip_node_name
                            End If
                        End If
                        If xmlNode.nodeName <> "TARGETFIELD" Then
skip_node_name:
                            Set newNode = xmlDom.createElement("TARGETFIELD")
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

                                    Set src_attr = xmlDom.createAttribute("KEYTYPE")
                                    src_attr.Value = port_keytype
                                    newNode.attributes.setNamedItem (src_attr)

                                    Set src_attr = xmlDom.createAttribute("NAME")
                                    src_attr.Value = port_name
                                    newNode.attributes.setNamedItem (src_attr)
                                    
                                    Set src_attr = xmlDom.createAttribute("NULLABLE")
                                    src_attr.Value = port_notnull
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
                                    
                            If xmlNode Is Nothing Then
                                Set xmlNode = xmlDom.selectSingleNode("//POWERMART/REPOSITORY/FOLDER/TARGET")
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
                            xmlNode.attributes.getNamedItem("SCALE").nodeValue = port_scale
                            
                            xmlNode.attributes.getNamedItem("KEYTYPE").nodeValue = port_keytype
                            xmlNode.attributes.getNamedItem("NULLABLE").nodeValue = port_notnull
                        End If
                        
                        Set xmlNode = xmlNode.nextSibling

                    Next output_at_row

                    'Remove port
                    If Not xmlNode Is Nothing Then
                        While xmlNode.nodeName = "TARGETFIELD"
                            Set xmlNode = xmlNode.nextSibling
                            xmlNode.previousSibling.parentNode.removeChild xmlNode.previousSibling
                        Wend
                    End If
            
            xmlDom.Save xml_filepath + "/" + xml_filename
            Set xmlNode = Nothing
            Set newNode = Nothing
            Set xmlTgtNode = Nothing
            Set xmlTgtNodeList = Nothing
            Call Sub_OkOnly_Msgbox("Complete update.")
            
            If mapping_select_file_flg = 1 And selected_trnsf_type = "Target Definition" Then
                ThisWorkbook.Sheets("edit_mapping").Activate
            End If
            Call Sub_Hint_Box_Add(Format(Time, "hh:mm:ss") + ": Port changes for " + tgt_name + " have been updated to the XML file." + vbLf)
            Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
            Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Edit_Tgt")
End Sub








