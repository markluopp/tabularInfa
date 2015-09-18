Attribute VB_Name = "Sorter"
'----------------------------------
'mluo@merkleinc.com
'Version:
'2015-5-18 intail version
'----------------------------------
Public Sub Sub_Edit_Srt(xmlDom As MSXML2.DOMDocument, srt_name As String)
On Error GoTo FATAL_ERROR
     Dim xmlNode As MSXML2.IXMLDOMNode
     Dim xmlNodeList As MSXML2.IXMLDOMNodeList
     Dim chlNode As MSXML2.IXMLDOMNode
     Dim chiNodeList As MSXML2.IXMLDOMNodeList
     
     If InStr(srt_name, "(") = 0 Then
        reuseable_flg = 0
        Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/MAPPING/TRANSFORMATION")
     Else
        reuseable_flg = 1
        srt_name = Mid(srt_name, InStr(srt_name, "(") + 1, Len(srt_name) - InStr(srt_name, "(") - 1)
        Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/TRANSFORMATION")
     End If
     
     output_at_row = 10
     '10th row is first row of output,detect max column number
     If [iv10].End(xlToLeft).Column > [iv9].End(xlToLeft).Column Then
        header_end_at = [iv10].End(xlToLeft).Column
     Else
        header_end_at = [iv9].End(xlToLeft).Column
     End If
     
     If header_end_at < 4 Then
        header_end_at = 4
     End If
     
     If ActiveSheet.Range("D65535").End(xlUp).row < 10 Then
        end_at_row = 10
     Else
        end_at_row = ActiveSheet.Range("D65535").End(xlUp).row
     End If

     ActiveSheet.Range("D" + CStr(end_at_row) + ":" + Chr(header_end_at + 64) + "10").Clear
     
     Set xmlNode = Nothing
     Set chlNodeList = Nothing

        For Each xmlNode In xmlNodeList
          If xmlNode.attributes.getNamedItem("NAME").nodeValue = srt_name Then
              Set chlNodeList = xmlNode.childNodes
                 For Each chlNode In chlNodeList
                    If chlNode.nodeName = "TRANSFORMFIELD" Then
                        port_name = chlNode.attributes.getNamedItem("NAME").nodeValue
                        port_datatype = chlNode.attributes.getNamedItem("DATATYPE").nodeValue
                        port_pre = chlNode.attributes.getNamedItem("PRECISION").nodeValue
                        port_scale = chlNode.attributes.getNamedItem("SCALE").nodeValue
                        is_key = chlNode.attributes.getNamedItem("ISSORTKEY").nodeValue
                        sort_direc = chlNode.attributes.getNamedItem("SORTDIRECTION").nodeValue
                            
                        ActiveSheet.Range("D" & output_at_row).Value = port_name
                        ActiveSheet.Range("E" & output_at_row).Value = port_datatype
                        ActiveSheet.Range("F" & output_at_row).Value = port_pre
                        ActiveSheet.Range("G" & output_at_row).Value = port_scale
                        ActiveSheet.Range("H" & output_at_row).Value = is_key
                        ActiveSheet.Range("I" & output_at_row).Value = sort_direc
                        
                        output_at_row = output_at_row + 1
                    End If
                Next
            End If
        Next
            
        If chlNodeList Is Nothing Then
            Call Sub_OkOnly_Msgbox("Can not find specified transformation '" + srt_name + "'.")
            Exit Sub
        End If
        
        ActiveSheet.Range("D" + CStr(ActiveSheet.Range("D65535").End(xlUp).row) + ":" + Chr(header_end_at + 64) + "9").Columns.AutoFit
        Set xmlNode = Nothing
        Set xmlNodeList = Nothing
        Set chlNode = Nothing
        Set chlNodeList = Nothing
        
        Call Sub_Hint_Box_Add(Format(Time, "hh:mm:ss") + ": You are editing " + srt_name + " and its port layout has displayed at right.You can modify these ports as you want, then click 'Update This Transformation' to save changes." + vbLf)
        Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
        
        Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Edit_Trnsf_Part")
End Sub

Public Sub Sub_Update_Srt(xmlDom As MSXML2.DOMDocument, srt_name As String)
On Error GoTo FATAL_ERROR
     Dim newNode As MSXML2.IXMLDOMNode
     Dim xmlNodeList As MSXML2.IXMLDOMNodeList
     Dim chlNode As MSXML2.IXMLDOMNode
     Dim fieldNode As MSXML2.IXMLDOMNode
     Dim fieldNodeList As MSXML2.IXMLDOMNodeList

            output_at_row = 10
            end_at_row = ActiveSheet.Range("D65535").End(xlUp).row

            If reuseable_flg = 0 Then
                Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/MAPPING/TRANSFORMATION")
            Else
                Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/TRANSFORMATION")
            End If
            Set xmlNode = Nothing
            For Each xmlNode In xmlNodeList
            
                Set fieldNodeList = xmlNode.selectNodes("TRANSFORMFIELD")
                If xmlNode.attributes.getNamedItem("NAME").nodeValue = srt_name Then
                
                    Set chlNode = xmlNode.FirstChild
                    
                    For output_at_row = 10 To end_at_row
                        
                        port_name = ActiveSheet.Range("D" & output_at_row).Value
                        For i = 10 To output_at_row
                            If port_name = ActiveSheet.Range("D" & i).Value And i <> output_at_row Then
                                ActiveSheet.Cells(i, "D").Interior.ColorIndex = 3
                                ActiveSheet.Cells(output_at_row, "D").Interior.ColorIndex = 3
                                Call Sub_OkOnly_Msgbox("Duplicated port name!")
                                Exit Sub
                            End If
                        Next
                        port_datatype = ActiveSheet.Range("E" & output_at_row).Value
                        port_pre = ActiveSheet.Range("F" & output_at_row).Value
                        port_scale = ActiveSheet.Range("G" & output_at_row).Value
                        is_key = ActiveSheet.Range("H" & output_at_row).Value
                        sort_direc = ActiveSheet.Range("I" & output_at_row).Value
                        
                        'Check invalid datatype
                        Select Case port_datatype
                        Case "bigint"
                            port_pre = 19
                            port_scale = 0
                        Case "date/time"
                            port_pre = 29
                            port_scale = 9
                        Case "double"
                            port_pre = 19
                            port_scale = 0
                        Case "integer"
                            port_pre = 10
                            port_scale = 0
                        Case "real"
                            port_pre = 7
                            port_scale = 0
                        Case "small integer"
                            port_pre = 5
                            port_scale = 0
                        Case "binary", "string", "nstring", "text", "ntext"
                            port_scale = 0
                        Case "decimal"
                        Case Else
                            ActiveSheet.Cells(output_at_row, "E").Interior.ColorIndex = 3
                            Call Sub_OkOnly_Msgbox("Invalid transformation data type '" + port_datatype + "' for informatica.")
                            Exit Sub
                        End Select
                        
                        'Check is key
                        Select Case is_key
                        Case "YES", "NO"
                        Case Else
                            ActiveSheet.Cells(output_at_row, "H").Interior.ColorIndex = 3
                            Call Sub_OkOnly_Msgbox("Please Only Input 'YES' Or 'NO'.")
                            Exit Sub
                        End Select
                        'Check sort_direc
                        Select Case sort_direc
                        Case "ASCENDING", "DESCENDING"
                        Case Else
                            ActiveSheet.Cells(output_at_row, "I").Interior.ColorIndex = 3
                            Call Sub_OkOnly_Msgbox("Please Only Input 'ASCENDING' Or 'DESCENDING'.")
                            Exit Sub
                        End Select
                        
                        'Add port
                        If chlNode.nodeName <> "TRANSFORMFIELD" Then
                            Set newNode = xmlDom.createElement("TRANSFORMFIELD")
                                    
                                    Set srt_attr = xmlDom.createAttribute("DATATYPE")
                                    srt_attr.Value = port_datatype
                                    newNode.attributes.setNamedItem (srt_attr)

                                    Set srt_attr = xmlDom.createAttribute("DEFAULTVALUE")
                                    srt_attr.Value = ""
                                    newNode.attributes.setNamedItem (srt_attr)

                                    Set srt_attr = xmlDom.createAttribute("DESCRIPTION")
                                    srt_attr.Value = ""
                                    newNode.attributes.setNamedItem (srt_attr)

                                    Set srt_attr = xmlDom.createAttribute("ISSORTKEY")
                                    srt_attr.Value = is_key
                                    newNode.attributes.setNamedItem (srt_attr)

                                    Set srt_attr = xmlDom.createAttribute("NAME")
                                    srt_attr.Value = port_name
                                    newNode.attributes.setNamedItem (srt_attr)

                                    Set srt_attr = xmlDom.createAttribute("PICTURETEXT")
                                    srt_attr.Value = ""
                                    newNode.attributes.setNamedItem (srt_attr)

                                    Set srt_attr = xmlDom.createAttribute("PORTTYPE")
                                    srt_attr.Value = "INPUT/OUTPUT"
                                    newNode.attributes.setNamedItem (srt_attr)

                                    Set srt_attr = xmlDom.createAttribute("PRECISION")
                                    srt_attr.Value = port_pre
                                    newNode.attributes.setNamedItem (srt_attr)

                                    Set srt_attr = xmlDom.createAttribute("SCALE")
                                    srt_attr.Value = port_scale
                                    newNode.attributes.setNamedItem (srt_attr)
                                    
                                    Set srt_attr = xmlDom.createAttribute("SORTDIRECTION")
                                    srt_attr.Value = sort_direc
                                    newNode.attributes.setNamedItem (srt_attr)
                            
                            chlNode.parentNode.insertBefore newNode, chlNode
                            Set chlNode = chlNode.previousSibling
                        
                        Else
                            chlNode.attributes.getNamedItem("NAME").nodeValue = port_name
                            chlNode.attributes.getNamedItem("DESCRIPTION").nodeValue = ""
                            chlNode.attributes.getNamedItem("DATATYPE").nodeValue = port_datatype
                            chlNode.attributes.getNamedItem("PRECISION").nodeValue = port_pre
                            chlNode.attributes.getNamedItem("SCALE").nodeValue = port_scale
                            chlNode.attributes.getNamedItem("ISSORTKEY").nodeValue = is_key
                            chlNode.attributes.getNamedItem("SORTDIRECTION").nodeValue = sort_direc
                            
                        End If
                            
                        Set chlNode = chlNode.nextSibling
                            
                    Next output_at_row
                    
                    'Remove port
                    While chlNode.nodeName = "TRANSFORMFIELD"
                        Set chlNode = chlNode.nextSibling
                        chlNode.previousSibling.parentNode.removeChild chlNode.previousSibling
                    Wend
                    
                End If
            Next
            
            xmlDom.Save xml_filepath + "\" + xml_filename
            Set xmlNodeList = Nothing
            Set xmlNode = Nothing
            Set chlNode = Nothing
            Set newNode = Nothing
            
            Call Sub_OkOnly_Msgbox("Complete update.")
            Call Sub_Hint_Box_Add(Format(Time, "hh:mm:ss") + ": Port changes for " + srt_name + " have been updated to the XML file." + vbLf)
            Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
        
            Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Update_Srt")
End Sub



