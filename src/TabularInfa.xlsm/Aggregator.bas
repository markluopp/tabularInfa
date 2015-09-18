Attribute VB_Name = "Aggregator"
'----------------------------------
'mluo@merkleinc.com
'Version:
'2015-6-28 intail version, similiar with Expression
'----------------------------------

Public Sub Sub_Edit_Agg(xmlDom As MSXML2.DOMDocument, agg_name As String)
On Error GoTo FATAL_ERROR
     Dim xmlNode As MSXML2.IXMLDOMNode
     Dim xmlNodeList As MSXML2.IXMLDOMNodeList
     Dim chlNode As MSXML2.IXMLDOMNode
     Dim chiNodeList As MSXML2.IXMLDOMNodeList
     
     If InStr(agg_name, "(") = 0 Then
        reuseable_flg = 0
        Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/MAPPING/TRANSFORMATION")
     Else
        reuseable_flg = 1
        agg_name = Mid(agg_name, InStr(agg_name, "(") + 1, Len(agg_name) - InStr(agg_name, "(") - 1)
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
          If xmlNode.attributes.getNamedItem("NAME").nodeValue = agg_name Then
              Set chlNodeList = xmlNode.childNodes
                 For Each chlNode In chlNodeList
                    If chlNode.nodeName = "TRANSFORMFIELD" Then
                        port_name = chlNode.attributes.getNamedItem("NAME").nodeValue
                        port_datatype = chlNode.attributes.getNamedItem("DATATYPE").nodeValue
                        port_pre = chlNode.attributes.getNamedItem("PRECISION").nodeValue
                        port_scale = chlNode.attributes.getNamedItem("SCALE").nodeValue
                        'Input port dont have expression
                        If Not chlNode.attributes.getNamedItem("EXPRESSION") Is Nothing Then
                            port_exp = chlNode.attributes.getNamedItem("EXPRESSION").nodeValue
                        Else
                            port_exp = ""
                        End If
                        port_type = chlNode.attributes.getNamedItem("PORTTYPE").nodeValue
                        port_exp_type = chlNode.attributes.getNamedItem("EXPRESSIONTYPE").nodeValue
                            
                        'fisrt single quote would hide in excel
                        If Right(port_exp, 1) = "'" Then
                            port_exp = "'" + port_exp
                        End If
                            
                        ActiveSheet.Range("D" & output_at_row).Value = port_name
                        ActiveSheet.Range("E" & output_at_row).Value = port_datatype
                        ActiveSheet.Range("F" & output_at_row).Value = port_pre
                        ActiveSheet.Range("G" & output_at_row).Value = port_scale
                        ActiveSheet.Range("H" & output_at_row).Value = port_exp
                        ActiveSheet.Range("I" & output_at_row).Value = port_type
                        ActiveSheet.Range("J" & output_at_row).Value = port_exp_type
                        output_at_row = output_at_row + 1
                    End If
                Next
            End If
        Next
            
        If chlNodeList Is Nothing Then
            Call Sub_OkOnly_Msgbox("Can not find specified expression '" + agg_name + "'.")
            Exit Sub
        End If
        
        ActiveSheet.Range("D" + CStr(ActiveSheet.Range("D65535").End(xlUp).row) + ":" + Chr(header_end_at + 64) + "9").Columns.AutoFit
        Set xmlNode = Nothing
        Set xmlNodeList = Nothing
        Set chlNode = Nothing
        Set chlNodeList = Nothing
        
        Call Sub_Hint_Box_Add(Format(Time, "hh:mm:ss") + ": You are editing " + agg_name + " and its port layout has displayed at right.You can modify these ports as you want, then click 'Update This Transformation' to save changes." + vbLf)
        Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
        
        Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Edit_Agg")
End Sub
'----------------------------------
'mluo@merkleinc.com
'Version:
'2015-4-12 intail version
'----------------------------------
Public Sub Sub_Update_Agg(xmlDom As MSXML2.DOMDocument, agg_name As String)
On Error GoTo FATAL_ERROR
     Dim newNode As MSXML2.IXMLDOMNode
     Dim xmlNodeList As MSXML2.IXMLDOMNodeList
     Dim chlNode As MSXML2.IXMLDOMNode
     Dim fieldNode As MSXML2.IXMLDOMNode
     Dim fieldNodeList As MSXML2.IXMLDOMNodeList

     comment_check_flg = 1

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
                'Set fieldNode = Nothing
                'check comment mark '--' or '//'
                If comment_check_flg = 1 Then
                    For Each fieldNode In fieldNodeList
                        If Not fieldNode.attributes.getNamedItem("EXPRESSION") Is Nothing Then
                            If InStr(1, fieldNode.attributes.getNamedItem("EXPRESSION").nodeValue, "--") <> 0 Or InStr(1, fieldNode.attributes.getNamedItem("EXPRESSION").nodeValue, "//") <> 0 Then
                                If MsgBox("Exist comment mark in expression " + xmlNode.attributes.getNamedItem("NAME").nodeValue + ". Click 'Yes' to stop checking(PLEASE DO REPLACE OPERATION BEFORE IMPORT!!). Click 'No' to end update.", vbYesNo) = vbNo Then
                                    Exit Sub
                                Else
                                    comment_check_flg = 0
                                End If
                            End If
                        End If
                    Next
                End If
                If xmlNode.attributes.getNamedItem("NAME").nodeValue = agg_name Then
                
                    Set chlNode = xmlNode.FirstChild
                    
                    For output_at_row = 10 To end_at_row
                        'check duplicated port name
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
                        port_exp = ActiveSheet.Range("H" & output_at_row).Value
                        port_type = ActiveSheet.Range("I" & output_at_row).Value
                        port_exp_type = ActiveSheet.Range("J" & output_at_row).Value
                        
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
                        Case "datetime"
                            ActiveSheet.Range("E" & output_at_row).Value = "date/time"
                            port_datatype = "date/time"
                            port_pre = 29
                            port_scale = 9
                        Case "int"
                            ActiveSheet.Range("E" & output_at_row).Value = "integer"
                            port_datatype = "integer"
                            port_pre = 29
                            port_scale = 9
                            port_pre = 10
                            port_scale = 0
                        Case Else
                            ActiveSheet.Cells(output_at_row, "E").Interior.ColorIndex = 3
                            Call Sub_OkOnly_Msgbox("Invalid transformation data type '" + port_datatype + "' for informatica.")
                            Exit Sub
                        End Select
                        
                        'Check port type
                        Select Case port_type
                        Case "INPUT"
                            If port_exp <> "" Then
                                ActiveSheet.Cells(output_at_row, "H").Interior.ColorIndex = 3
                                Call Sub_OkOnly_Msgbox("Input port shouldn't have expression value!")
                                Exit Sub
                            End If
                        Case "INPUT/OUTPUT"
                            If port_exp <> port_name Then
                                ActiveSheet.Cells(output_at_row, "H").Interior.ColorIndex = 3
                                Call Sub_OkOnly_Msgbox("Input/output port shouldn't change expression value!")
                                Exit Sub
                            End If
                        Case "OUTPUT", "LOCAL VARIABLE"
                        Case Else
                            ActiveSheet.Cells(output_at_row, "I").Interior.ColorIndex = 3
                            Call Sub_OkOnly_Msgbox("Invalid transformation port type '" + port_type + "' for informatica.")
                            Exit Sub
                        End Select
                        
                        'Check group by
                        Select Case port_exp_type
                        Case "GROUPBY", "GENERAL"
                        Case Else
                            Call Sub_OkOnly_Msgbox("Please Only Input 'GENERAL' Or 'GROUPBY'.")
                            Exit Sub
                        End Select
                        
                        'Add port
                        If chlNode.nodeName <> "TRANSFORMFIELD" Then
                            Set newNode = xmlDom.createElement("TRANSFORMFIELD")
                                    
                                    Set exp_attr = xmlDom.createAttribute("DATATYPE")
                                    exp_attr.Value = port_datatype
                                    newNode.attributes.setNamedItem (exp_attr)

                                    Set exp_attr = xmlDom.createAttribute("DEFAULTVALUE")
                                    exp_attr.Value = ""
                                    newNode.attributes.setNamedItem (exp_attr)

                                    Set exp_attr = xmlDom.createAttribute("DESCRIPTION")
                                    exp_attr.Value = ""
                                    newNode.attributes.setNamedItem (exp_attr)

                                    Set exp_attr = xmlDom.createAttribute("EXPRESSION")
                                    exp_attr.Value = port_exp
                                    newNode.attributes.setNamedItem (exp_attr)

                                    Set exp_attr = xmlDom.createAttribute("EXPRESSIONTYPE")
                                    exp_attr.Value = port_exp_type
                                    newNode.attributes.setNamedItem (exp_attr)

                                    Set exp_attr = xmlDom.createAttribute("NAME")
                                    exp_attr.Value = port_name
                                    newNode.attributes.setNamedItem (exp_attr)

                                    Set exp_attr = xmlDom.createAttribute("PICTURETEXT")
                                    exp_attr.Value = ""
                                    newNode.attributes.setNamedItem (exp_attr)

                                    Set exp_attr = xmlDom.createAttribute("PORTTYPE")
                                    exp_attr.Value = port_type
                                    newNode.attributes.setNamedItem (exp_attr)

                                    Set exp_attr = xmlDom.createAttribute("PRECISION")
                                    exp_attr.Value = port_pre
                                    newNode.attributes.setNamedItem (exp_attr)

                                    Set exp_attr = xmlDom.createAttribute("SCALE")
                                    exp_attr.Value = port_scale
                                    newNode.attributes.setNamedItem (exp_attr)
                            'input port node lack two attributes(expression and expressiontype)
                            If port_type = "INPUT" Then
                                    newNode.attributes.removeNamedItem ("EXPRESSION")
                                    newNode.attributes.removeNamedItem ("EXPRESSIONTYPE")
                            End If
                            
                            chlNode.parentNode.insertBefore newNode, chlNode
                            Set chlNode = chlNode.previousSibling
                        
                        Else
                            chlNode.attributes.getNamedItem("NAME").nodeValue = port_name
                            chlNode.attributes.getNamedItem("DESCRIPTION").nodeValue = ""
                            chlNode.attributes.getNamedItem("DATATYPE").nodeValue = port_datatype
                            chlNode.attributes.getNamedItem("PRECISION").nodeValue = port_pre
                            chlNode.attributes.getNamedItem("SCALE").nodeValue = port_scale
                            chlNode.attributes.getNamedItem("PORTTYPE").nodeValue = port_type
                            chlNode.attributes.getNamedItem("EXPRESSIONTYPE").nodeValue = port_exp_type
                            
                            If port_type <> "INPUT" Then
                                If chlNode.attributes.getNamedItem("EXPRESSION") Is Nothing Then
                                'MsgBox output_at_row
                                    Set exp_attr = xmlDom.createAttribute("EXPRESSION")
                                    exp_attr.Value = ""
                                    chlNode.attributes.setNamedItem (exp_attr)
                                    Set exptype_attr = xmlDom.createAttribute("EXPRESSIONTYPE")
                                    exptype_attr.Value = "GENERAL"
                                    chlNode.attributes.setNamedItem (exptype_attr)
                                End If
                                chlNode.attributes.getNamedItem("EXPRESSION").nodeValue = port_exp
                            Else
                                If Not chlNode.attributes.getNamedItem("EXPRESSION") Is Nothing Then
                                    chlNode.attributes.removeNamedItem ("EXPRESSION")
                                    chlNode.attributes.removeNamedItem ("EXPRESSIONTYPE")
                                End If
                            End If
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
            
            Call Sub_Hint_Box_Add(Format(Time, "hh:mm:ss") + ": Port changes for " + agg_name + " have been updated to the XML file." + vbLf)
            Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
        
            Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Update_Agg")
End Sub




