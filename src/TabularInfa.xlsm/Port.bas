Attribute VB_Name = "Port"
'----------------------------------
'mluo@merkleinc.com
'Version:
'2015-5-19 intail version
'2015-5-25 add link according to existent links
'----------------------------------
'Public propagate_staus As String
'0-----ready
'1-----select target transformation
Public Sub Sub_Propagate_Port(fix_type, fix_value, xmlDom As MSXML2.DOMDocument, start_name, start_type, trnsf_name() As String, trnsf_type() As String, port_name() As String, port_data_type() As String, port_prec() As String, port_scale() As String)
'On Error GoTo FATAL_ERROR
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim xmlNodeList As MSXML2.IXMLDOMNodeList
    Dim locateNode As MSXML2.IXMLDOMNode
    Dim chiNodeList As MSXML2.IXMLDOMNodeList
    null_trnsf_flg = 0

    Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/MAPPING/TRANSFORMATION")
    'traverse all tranformation name
    For i = 1 To UBound(trnsf_name)
        For Each xmlNode In xmlNodeList
            If xmlNode.attributes.getNamedItem("NAME").nodeValue = trnsf_name(i - 1) Then
                'MsgBox "find " + trnsf_name(i - 1)
                'locate output node
                Set chiNodeList = xmlNode.selectNodes("TRANSFORMFIELD")
                Set locateNode = chiNodeList.Item(chiNodeList.Length - 1)
                If locateNode Is Nothing Then
                    Set locateNode = xmlNode.FirstChild
                    null_trnsf_flg = 1
                End If
'                If locateNode Is Nothing Then
'                MsgBox "cannot find"
'                End If
'Custom Transformation has TEMPLATENAME attribute
                If xmlNode.attributes.getNamedItem("TYPE").nodeValue = "Custom Transformation" Then
                    Select Case xmlNode.attributes.getNamedItem("TEMPLATENAME").nodeValue
                    Case "Union Transformation"
                        cur_node_trnsf_type = xmlNode.attributes.getNamedItem("TEMPLATENAME").nodeValue
                    Case Else
                        Call Sub_OkOnly_Msgbox("So Far, Only Support Union In Custom Transformations.")
                        Exit Sub
                    End Select
                Else
                    cur_node_trnsf_type = xmlNode.attributes.getNamedItem("TYPE").nodeValue
                End If
                Select Case cur_node_trnsf_type
                'add different port type according to transformation type
                'Similar Transformation Field Type
                Case "Expression", "Aggregator"
                    For j = 1 To UBound(port_name)
                    If Func_Existent_Check(xmlNode, port_name(j - 1)) = 0 Then
                                    Set newNode = xmlDom.createElement("TRANSFORMFIELD")
                                    
                                    Set exp_attr = xmlDom.createAttribute("DATATYPE")
                                    exp_attr.Value = port_data_type(j - 1)
                                    newNode.attributes.setNamedItem (exp_attr)

                                    Set exp_attr = xmlDom.createAttribute("DEFAULTVALUE")
                                    exp_attr.Value = ""
                                    newNode.attributes.setNamedItem (exp_attr)

                                    Set exp_attr = xmlDom.createAttribute("DESCRIPTION")
                                    exp_attr.Value = ""
                                    newNode.attributes.setNamedItem (exp_attr)

                                    Set exp_attr = xmlDom.createAttribute("EXPRESSION")
                                    exp_attr.Value = port_name(j - 1)
                                    newNode.attributes.setNamedItem (exp_attr)

                                    Set exp_attr = xmlDom.createAttribute("EXPRESSIONTYPE")
                                    exp_attr.Value = "GENERAL"
                                    newNode.attributes.setNamedItem (exp_attr)

                                    Set exp_attr = xmlDom.createAttribute("NAME")
                                    exp_attr.Value = port_name(j - 1)
                                    newNode.attributes.setNamedItem (exp_attr)

                                    Set exp_attr = xmlDom.createAttribute("PICTURETEXT")
                                    exp_attr.Value = ""
                                    newNode.attributes.setNamedItem (exp_attr)

                                    Set exp_attr = xmlDom.createAttribute("PORTTYPE")
                                    exp_attr.Value = "INPUT/OUTPUT"
                                    newNode.attributes.setNamedItem (exp_attr)

                                    Set exp_attr = xmlDom.createAttribute("PRECISION")
                                    exp_attr.Value = port_prec(j - 1)
                                    newNode.attributes.setNamedItem (exp_attr)

                                    Set exp_attr = xmlDom.createAttribute("SCALE")
                                    exp_attr.Value = port_scale(j - 1)
                                    newNode.attributes.setNamedItem (exp_attr)
                            If null_trnsf_flg = 0 Then
                                locateNode.parentNode.insertBefore newNode, locateNode.nextSibling
                            Else
                                locateNode.parentNode.insertBefore newNode, locateNode
                            End If
                        Else
                            If MsgBox("Port " + port_name(j - 1) + " Has Already Exist In " + trnsf_name(i - 1) + "! Click 'Yes' To Skip This Port And Continue Next. Click 'No' To Rollback Propagation.", vbYesNo, "TabularInfa") = vbNo Then
                                Exit Sub
                            End If
                        End If
                    Next
                Case "Sorter"
                    For j = 1 To UBound(port_name)
                    If Func_Existent_Check(xmlNode, port_name(j - 1)) = 0 Then
                                    Set newNode = xmlDom.createElement("TRANSFORMFIELD")
                                    
                                    Set srt_attr = xmlDom.createAttribute("DATATYPE")
                                    srt_attr.Value = port_data_type(j - 1)
                                    newNode.attributes.setNamedItem (srt_attr)

                                    Set srt_attr = xmlDom.createAttribute("DEFAULTVALUE")
                                    srt_attr.Value = ""
                                    newNode.attributes.setNamedItem (srt_attr)

                                    Set srt_attr = xmlDom.createAttribute("DESCRIPTION")
                                    srt_attr.Value = "INPUT"
                                    newNode.attributes.setNamedItem (srt_attr)

                                    Set srt_attr = xmlDom.createAttribute("ISSORTKEY")
                                    srt_attr.Value = "NO"
                                    newNode.attributes.setNamedItem (srt_attr)

                                    Set srt_attr = xmlDom.createAttribute("NAME")
                                    srt_attr.Value = port_name(j - 1)
                                    newNode.attributes.setNamedItem (srt_attr)

                                    Set srt_attr = xmlDom.createAttribute("PICTURETEXT")
                                    srt_attr.Value = ""
                                    newNode.attributes.setNamedItem (srt_attr)

                                    Set srt_attr = xmlDom.createAttribute("PORTTYPE")
                                    srt_attr.Value = "INPUT/OUTPUT"
                                    newNode.attributes.setNamedItem (srt_attr)

                                    Set srt_attr = xmlDom.createAttribute("PRECISION")
                                    srt_attr.Value = port_prec(j - 1)
                                    newNode.attributes.setNamedItem (srt_attr)

                                    Set srt_attr = xmlDom.createAttribute("SCALE")
                                    srt_attr.Value = port_scale(j - 1)
                                    newNode.attributes.setNamedItem (srt_attr)
                                    
                                    Set srt_attr = xmlDom.createAttribute("SORTDIRECTION")
                                    srt_attr.Value = "ASCENDING"
                                    newNode.attributes.setNamedItem (srt_attr)
                           
                            If null_trnsf_flg = 0 Then
                                locateNode.parentNode.insertBefore newNode, locateNode.nextSibling
                            Else
                                locateNode.parentNode.insertBefore newNode, locateNode
                            End If
                        Else
                            If MsgBox("Port " + port_name(j - 1) + " Has Already Exist In " + trnsf_name(i - 1) + "! Click 'Yes' To Skip This Port And Continue Next. Click 'No' To Rollback Propagation.", vbYesNo, "TabularInfa") = vbNo Then
                                Exit Sub
                            End If
                        End If
                    Next
                    
                Case "Router", "Union Transformation"
                'Set the suffix number for new added port
                output_count = 0
                    For Each trnsfNode In chiNodeList
                        'Check if it's last node
                        If Not trnsfNode.nextSibling Is Nothing Then
                            'Skip TABLEATTRIBUTE node
                            If trnsfNode.nextSibling.nodeName = "TRANSFORMFIELD" Then
                                If trnsfNode.attributes.getNamedItem("GROUP").nodeValue <> trnsfNode.nextSibling.attributes.getNamedItem("GROUP").nodeValue Then
                                    'trnsfNode.attributes.getNamedItem("NAME").nodeValue <> trnsfNode.nextSibling.attributes.getNamedItem("NAME").nodeValue
                                    Set locateNode = trnsfNode
                                    If cur_node_trnsf_type = "Union Transformation" Then
                                        group_check_value = "OUTPUT"
                                    Else
                                        group_check_value = "INPUT"
                                    End If
                                    'Add port to input group
                                    If trnsfNode.attributes.getNamedItem("GROUP").nodeValue = group_check_value Then
                                        For j = 1 To UBound(port_name)
                                            If Func_Existent_Check(xmlNode, port_name(j - 1)) = 0 Then
                                                Set newNode = xmlDom.createElement("TRANSFORMFIELD")

                                                Set rtr_attr = xmlDom.createAttribute("DATATYPE")
                                                rtr_attr.Value = port_data_type(j - 1)
                                                newNode.attributes.setNamedItem (rtr_attr)

                                                Set rtr_attr = xmlDom.createAttribute("DEFAULTVALUE")
                                                rtr_attr.Value = ""
                                                newNode.attributes.setNamedItem (rtr_attr)

                                                Set rtr_attr = xmlDom.createAttribute("DESCRIPTION")
                                                rtr_attr.Value = ""
                                                newNode.attributes.setNamedItem (rtr_attr)

                                                Set rtr_attr = xmlDom.createAttribute("GROUP")
                                                If cur_node_trnsf_type = "Union Transformation" Then
                                                    rtr_attr.Value = "OUTPUT"
                                                Else
                                                    rtr_attr.Value = "INPUT"
                                                End If
                                                newNode.attributes.setNamedItem (rtr_attr)

                                                Set rtr_attr = xmlDom.createAttribute("NAME")
                                                rtr_attr.Value = port_name(j - 1)
                                                newNode.attributes.setNamedItem (rtr_attr)
                                                
                                                If cur_node_trnsf_type = "Union Transformation" Then
                                                    Set rtr_attr = xmlDom.createAttribute("OUTPUTGROUP")
                                                    rtr_attr.Value = "OUTPUT"
                                                    newNode.attributes.setNamedItem (rtr_attr)
                                                End If

                                                Set rtr_attr = xmlDom.createAttribute("PICTURETEXT")
                                                rtr_attr.Value = ""
                                                newNode.attributes.setNamedItem (rtr_attr)

                                                Set rtr_attr = xmlDom.createAttribute("PORTTYPE")
                                                If cur_node_trnsf_type = "Union Transformation" Then
                                                    rtr_attr.Value = "OUTPUT"
                                                Else
                                                    rtr_attr.Value = "INPUT"
                                                End If
                                                newNode.attributes.setNamedItem (rtr_attr)

                                                Set rtr_attr = xmlDom.createAttribute("PRECISION")
                                                rtr_attr.Value = port_prec(j - 1)
                                                newNode.attributes.setNamedItem (rtr_attr)

                                                Set rtr_attr = xmlDom.createAttribute("SCALE")
                                                rtr_attr.Value = port_scale(j - 1)
                                                newNode.attributes.setNamedItem (rtr_attr)

                                                If null_trnsf_flg = 0 Then
                                                    locateNode.parentNode.insertBefore newNode, locateNode.nextSibling
                                                Else
                                                    locateNode.parentNode.insertBefore newNode, locateNode
                                                End If
                                            Else
                                                If MsgBox("Port " + port_name(j - 1) + " Has Already Exist In " + trnsf_name(i - 1) + "! Click 'Yes' To Skip This Port And Continue Next. Click 'No' To Rollback Propagation.", vbYesNo, "TabularInfa") = vbNo Then
                                                    Exit Sub
                                                End If
                                            End If
                                        Next
                                    Else
                                    'Add port to other output group
                                        output_count = output_count + 1
                                        For j = 1 To UBound(port_name)

                                            Set newNode = xmlDom.createElement("TRANSFORMFIELD")

                                            Set rtr_attr = xmlDom.createAttribute("DATATYPE")
                                            rtr_attr.Value = port_data_type(j - 1)
                                            newNode.attributes.setNamedItem (rtr_attr)

                                            Set rtr_attr = xmlDom.createAttribute("DEFAULTVALUE")
                                            rtr_attr.Value = ""
                                            newNode.attributes.setNamedItem (rtr_attr)

                                            Set rtr_attr = xmlDom.createAttribute("DESCRIPTION")
                                            rtr_attr.Value = ""
                                            newNode.attributes.setNamedItem (rtr_attr)

                                            Set rtr_attr = xmlDom.createAttribute("GROUP")
                                            rtr_attr.Value = trnsfNode.attributes.getNamedItem("GROUP").nodeValue
                                            newNode.attributes.setNamedItem (rtr_attr)

                                            Set rtr_attr = xmlDom.createAttribute("NAME")
                                            rtr_attr.Value = port_name(j - 1) + CStr(output_count)
                                            newNode.attributes.setNamedItem (rtr_attr)
                                            
                                            If cur_node_trnsf_type = "Union Transformation" Then
                                                Set rtr_attr = xmlDom.createAttribute("OUTPUTGROUP")
                                                rtr_attr.Value = trnsfNode.attributes.getNamedItem("GROUP").nodeValue
                                                newNode.attributes.setNamedItem (rtr_attr)
                                            End If

                                            Set rtr_attr = xmlDom.createAttribute("PICTURETEXT")
                                            rtr_attr.Value = ""
                                            newNode.attributes.setNamedItem (rtr_attr)

                                            Set rtr_attr = xmlDom.createAttribute("PORTTYPE")
                                            If cur_node_trnsf_type = "Union Transformation" Then
                                                rtr_attr.Value = "INPUT"
                                            Else
                                                rtr_attr.Value = "OUTPUT"
                                            End If
                                            newNode.attributes.setNamedItem (rtr_attr)

                                            Set rtr_attr = xmlDom.createAttribute("PRECISION")
                                            rtr_attr.Value = port_prec(j - 1)
                                            newNode.attributes.setNamedItem (rtr_attr)

                                            Set rtr_attr = xmlDom.createAttribute("REF_FIELD")
                                            rtr_attr.Value = port_name(j - 1)
                                            newNode.attributes.setNamedItem (rtr_attr)

                                            Set rtr_attr = xmlDom.createAttribute("SCALE")
                                            rtr_attr.Value = port_scale(j - 1)
                                            newNode.attributes.setNamedItem (rtr_attr)

                                            If null_trnsf_flg = 0 Then
                                                locateNode.parentNode.insertBefore newNode, locateNode.nextSibling
                                            Else
                                                locateNode.parentNode.insertBefore newNode, locateNode
                                            End If
                                            
                                            'Add FIELDDEPENDENCY Node for Union Transformation
                                            If cur_node_trnsf_type = "Union Transformation" Then
                                                Set newNode = xmlDom.createElement("FIELDDEPENDENCY")
    
                                                Set rtr_attr = xmlDom.createAttribute("INPUTFIELD")
                                                rtr_attr.Value = port_name(j - 1) + CStr(output_count)
                                                newNode.attributes.setNamedItem (rtr_attr)
    
                                                Set rtr_attr = xmlDom.createAttribute("OUTPUTFIELD")
                                                rtr_attr.Value = port_name(j - 1)
                                                newNode.attributes.setNamedItem (rtr_attr)
                                                
                                                xmlNode.appendChild newNode
                                            End If
                                        Next
                                    End If
                                End If
                            Else
                            'Add port to last group
                            Set locateNode = trnsfNode
                            output_count = output_count + 1
                                 For j = 1 To UBound(port_name)
                                        Set newNode = xmlDom.createElement("TRANSFORMFIELD")
    
                                        Set rtr_attr = xmlDom.createAttribute("DATATYPE")
                                        rtr_attr.Value = port_data_type(j - 1)
                                        newNode.attributes.setNamedItem (rtr_attr)
    
                                        Set rtr_attr = xmlDom.createAttribute("DEFAULTVALUE")
                                        rtr_attr.Value = ""
                                        newNode.attributes.setNamedItem (rtr_attr)
    
                                        Set rtr_attr = xmlDom.createAttribute("DESCRIPTION")
                                        rtr_attr.Value = ""
                                        newNode.attributes.setNamedItem (rtr_attr)
    
                                        Set rtr_attr = xmlDom.createAttribute("GROUP")
                                        rtr_attr.Value = trnsfNode.attributes.getNamedItem("GROUP").nodeValue
                                        newNode.attributes.setNamedItem (rtr_attr)
                                        
                                        Set rtr_attr = xmlDom.createAttribute("NAME")
                                        rtr_attr.Value = port_name(j - 1) + CStr(output_count)
                                        newNode.attributes.setNamedItem (rtr_attr)
                                        
                                        If cur_node_trnsf_type = "Union Transformation" Then
                                            Set rtr_attr = xmlDom.createAttribute("OUTPUTGROUP")
                                            rtr_attr.Value = trnsfNode.attributes.getNamedItem("GROUP").nodeValue
                                            newNode.attributes.setNamedItem (rtr_attr)
                                        End If
                                            
                                        Set rtr_attr = xmlDom.createAttribute("PICTURETEXT")
                                        rtr_attr.Value = ""
                                        newNode.attributes.setNamedItem (rtr_attr)
    
                                        Set rtr_attr = xmlDom.createAttribute("PORTTYPE")
                                        If cur_node_trnsf_type = "Union Transformation" Then
                                            rtr_attr.Value = "INPUT"
                                        Else
                                            rtr_attr.Value = "OUTPUT"
                                        End If
                                        newNode.attributes.setNamedItem (rtr_attr)
    
                                        Set rtr_attr = xmlDom.createAttribute("PRECISION")
                                        rtr_attr.Value = port_prec(j - 1)
                                        newNode.attributes.setNamedItem (rtr_attr)
                                        
                                        Set rtr_attr = xmlDom.createAttribute("REF_FIELD")
                                        rtr_attr.Value = port_name(j - 1)
                                        newNode.attributes.setNamedItem (rtr_attr)
                                        
                                        Set rtr_attr = xmlDom.createAttribute("SCALE")
                                        rtr_attr.Value = port_scale(j - 1)
                                        newNode.attributes.setNamedItem (rtr_attr)
    
                                        If null_trnsf_flg = 0 Then
                                            locateNode.parentNode.insertBefore newNode, locateNode.nextSibling
                                        Else
                                            locateNode.parentNode.insertBefore newNode, locateNode
                                        End If
                                        
                                        'Add FIELDDEPENDENCY Node for Union Transformation
                                        If cur_node_trnsf_type = "Union Transformation" Then
                                            Set newNode = xmlDom.createElement("FIELDDEPENDENCY")
    
                                            Set rtr_attr = xmlDom.createAttribute("INPUTFIELD")
                                            rtr_attr.Value = port_name(j - 1) + CStr(output_count)
                                            newNode.attributes.setNamedItem (rtr_attr)
    
                                            Set rtr_attr = xmlDom.createAttribute("OUTPUTFIELD")
                                            rtr_attr.Value = port_name(j - 1)
                                            newNode.attributes.setNamedItem (rtr_attr)
                                                
                                            xmlNode.appendChild newNode
                                        End If
                                Next
                            End If
                        End If
                    Next
                    
                'General Transformation Field Type
                Case "Joiner", "Source Qualifier", "Filter"
                    If xmlNode.attributes.getNamedItem("TYPE").nodeValue = "Joiner" Then
                        If MsgBox("Find Joiner " + trnsf_name(i - 1) + ", Yes to propagate Master port, No to propagate Detail port.", vbYesNo, "TabularInfa") = vbYes Then
                            port_type = "INPUT/OUTPUT/MASTER"
                        Else
                            port_type = "INPUT/OUTPUT"
                        End If
                    Else
                        port_type = "INPUT/OUTPUT"
                    End If
                    For j = 1 To UBound(port_name)
                    If Func_Existent_Check(xmlNode, port_name(j - 1)) = 0 Then
                                    Set newNode = xmlDom.createElement("TRANSFORMFIELD")
                                    
                                    Set trnsf_field_attr = xmlDom.createAttribute("DATATYPE")
                                    trnsf_field_attr.Value = port_data_type(j - 1)
                                    newNode.attributes.setNamedItem (trnsf_field_attr)

                                    Set trnsf_field_attr = xmlDom.createAttribute("DEFAULTVALUE")
                                    trnsf_field_attr.Value = ""
                                    newNode.attributes.setNamedItem (trnsf_field_attr)

                                    Set trnsf_field_attr = xmlDom.createAttribute("DESCRIPTION")
                                    trnsf_field_attr.Value = ""
                                    newNode.attributes.setNamedItem (trnsf_field_attr)

                                    Set trnsf_field_attr = xmlDom.createAttribute("NAME")
                                    trnsf_field_attr.Value = port_name(j - 1)
                                    newNode.attributes.setNamedItem (trnsf_field_attr)

                                    Set trnsf_field_attr = xmlDom.createAttribute("PICTURETEXT")
                                    trnsf_field_attr.Value = ""
                                    newNode.attributes.setNamedItem (trnsf_field_attr)

                                    Set trnsf_field_attr = xmlDom.createAttribute("PORTTYPE")
                                    trnsf_field_attr.Value = port_type
                                    newNode.attributes.setNamedItem (trnsf_field_attr)

                                    Set trnsf_field_attr = xmlDom.createAttribute("PRECISION")
                                    trnsf_field_attr.Value = port_prec(j - 1)
                                    newNode.attributes.setNamedItem (trnsf_field_attr)

                                    Set trnsf_field_attr = xmlDom.createAttribute("SCALE")
                                    trnsf_field_attr.Value = port_scale(j - 1)
                                    newNode.attributes.setNamedItem (trnsf_field_attr)
                           
                            If null_trnsf_flg = 0 Then
                                locateNode.parentNode.insertBefore newNode, locateNode.nextSibling
                            Else
                                locateNode.parentNode.insertBefore newNode, locateNode
                            End If
                        Else
                            If MsgBox("Port " + port_name(j - 1) + " Has Already Exist In " + trnsf_name(i - 1) + "! Click 'Yes' To Skip This Port And Continue Next. Click 'No' To Rollback Propagation.", vbYesNo, "TabularInfa") = vbNo Then
                                Exit Sub
                            End If
                        End If
                    Next
                
                End Select
            End If
        Next
    Next
    
    If MsgBox("Do You Want To Create Links For These Added Ports According To Existent Links?", vbYesNo, "TabularInfa") = vbYes Then
        'add start trasformation
        ReDim Preserve trnsf_name(UBound(trnsf_name) + 1)
        ReDim Preserve trnsf_type(UBound(trnsf_type) + 1)
        If InStr(start_name, "(") = 0 Then
            trnsf_name(UBound(trnsf_name) - 1) = start_name
        Else
            trnsf_name(UBound(trnsf_name) - 1) = Mid(start_name, 1, InStr(start_name, "(") - 1)
        End If
        If InStr(start_type, "(") = 0 Then
            trnsf_type(UBound(trnsf_type) - 1) = start_type
        Else
            trnsf_type(UBound(trnsf_type) - 1) = Mid(start_type, 1, InStr(start_type, "(") - 1)
        End If
        
'        MsgBox trnsf_name(UBound(trnsf_name))
'        MsgBox trnsf_type(UBound(trnsf_type))
'        Exit Sub
        'Check existent link
        Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/MAPPING/CONNECTOR")
        For i = 1 To UBound(trnsf_name)
            For j = 1 To UBound(trnsf_name)
                For Each xmlNode In xmlNodeList
                    If xmlNode.attributes.getNamedItem("FROMINSTANCE").nodeValue = trnsf_name(i - 1) And xmlNode.attributes.getNamedItem("TOINSTANCE").nodeValue = trnsf_name(j - 1) Then
                        For k = 1 To UBound(port_name)
                            'create link node
                            Set newNode = xmlDom.createElement("CONNECTOR")
                            
                            Set con_attr = xmlDom.createAttribute("FROMFIELD")
                            'remove here
                            If i = UBound(trnsf_name) Then
                                Select Case fix_type
                                Case "prefix"
                                    con_attr.Value = Mid(port_name(k - 1), Len(fix_value) + 1, Len(port_name(k - 1)) - Len(fix_value))
                                Case "suffix"
                                    con_attr.Value = Mid(port_name(k - 1), 1, Len(port_name(k - 1)) - Len(fix_value))
                                Case ""
                                    con_attr.Value = port_name(k - 1)
                                End Select
                            Else
                                con_attr.Value = port_name(k - 1)
                            End If
                            
                            'Router output port has number suffix
                            If xmlNode.attributes.getNamedItem("FROMINSTANCETYPE").nodeValue = "Router" Then
                                con_attr.Value = con_attr.Value + Right(xmlNode.attributes.getNamedItem("FROMFIELD").nodeValue, 1)
                            End If

                            newNode.attributes.setNamedItem (con_attr)
                            
                            Set con_attr = xmlDom.createAttribute("FROMINSTANCE")
                            con_attr.Value = trnsf_name(i - 1)
                            newNode.attributes.setNamedItem (con_attr)
                            
                            Set con_attr = xmlDom.createAttribute("FROMINSTANCETYPE")
                            con_attr.Value = trnsf_type(i - 1)
                            newNode.attributes.setNamedItem (con_attr)
                            
                            Set con_attr = xmlDom.createAttribute("TOFIELD")
                            con_attr.Value = port_name(k - 1)
                            'Only support "Union Transformation" in "Custom Transformation"
                            If xmlNode.attributes.getNamedItem("TOINSTANCETYPE").nodeValue = "Custom Transformation" Then
                                'And LCase(Left(xmlNode.attributes.getNamedItem("TOINSTANCE").nodeValue, 3)) = "uni"
                                con_attr.Value = con_attr.Value + Right(xmlNode.attributes.getNamedItem("TOFIELD").nodeValue, 1)
                            End If
                            newNode.attributes.setNamedItem (con_attr)
                            
                            Set con_attr = xmlDom.createAttribute("TOINSTANCE")
                            con_attr.Value = trnsf_name(j - 1)
                            newNode.attributes.setNamedItem (con_attr)
                            
                            Set con_attr = xmlDom.createAttribute("TOINSTANCETYPE")
                            con_attr.Value = trnsf_type(j - 1)
                            newNode.attributes.setNamedItem (con_attr)
                            
                            xmlNode.parentNode.insertBefore newNode, xmlNodeList.Item(xmlNodeList.Length - 1).nextSibling
                        Next
                        GoTo already_add_link_skip_to_next
                    End If
                Next
already_add_link_skip_to_next:
            Next
        Next
    End If
    
    xmlDom.Save xml_filepath + "\" + xml_filename
    
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set locateNode = Nothing
    Set chlNodeList = Nothing
    Set newNode = Nothing
            
    Call Sub_OkOnly_Msgbox("Complete Propagation.")
    
        Call Sub_Hint_Box_Add(Format(Time, "hh:mm:ss") + ": Propagation changes have been updated to the XML file." + vbLf)
        Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
        
        Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Propagate_Port")
End Sub

Public Function Func_Existent_Check(xmlNode As MSXML2.IXMLDOMNode, port_name As String) As Integer
    Dim checkNodeList As MSXML2.IXMLDOMNodeList
    Set checkNodeList = xmlNode.selectNodes("TRANSFORMFIELD")
    For Each trnsf_field In checkNodeList
        If trnsf_field.attributes.getNamedItem("NAME").nodeValue = port_name Then
            Func_Existent_Check = 1
            Exit Function
        End If
    Next
    Func_Existent_Check = 0
End Function

'----------------------------------
'Version:
'2015-6-25 Initial version, use to highlight backward and forward trnsf name
'---------------------------------
Public Sub Sub_Highlight(trnsf_name As String, xmlDom As MSXML2.DOMDocument)
On Error GoTo FATAL_ERROR
     Dim xmlNode As MSXML2.IXMLDOMNode
     Dim xmlNodeList As MSXML2.IXMLDOMNodeList
     Dim backward_input_link() As String
     Dim forward_input_link() As String
     
     If InStr(trnsf_name, "(") <> 0 Then
        trnsf_name = Mid(trnsf_name, 1, InStr(trnsf_name, "(") - 1)
     End If
     
     'max count is 101
     ReDim backward_input_link(100)
     ReDim forward_input_link(100)
     i = 0
     j = 0
     exist_flg = 0
     
     Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/MAPPING/CONNECTOR")

        For Each xmlNode In xmlNodeList
           If xmlNode.attributes.getNamedItem("TOINSTANCE").nodeValue = trnsf_name Then
                For i_check = 0 To i - 1
                    If xmlNode.attributes.getNamedItem("FROMINSTANCE").nodeValue = backward_input_link(i_check) Then
                        exist_flg = 1
                    End If
                Next
                If exist_flg = 0 Then
                    backward_input_link(i) = xmlNode.attributes.getNamedItem("FROMINSTANCE").nodeValue
                    i = i + 1
                End If
                exist_flg = 0
           End If
           If xmlNode.attributes.getNamedItem("FROMINSTANCE").nodeValue = trnsf_name Then
                'check if exist this trnsf name
                For j_check = 0 To j - 1
                    If xmlNode.attributes.getNamedItem("TOINSTANCE").nodeValue = forward_input_link(j_check) Then
                        exist_flg = 1
                    End If
                Next
                If exist_flg = 0 Then
                    forward_input_link(j) = xmlNode.attributes.getNamedItem("TOINSTANCE").nodeValue
                    j = j + 1
                End If
                exist_flg = 0
           End If
        Next
    If ConsoleForm.Console_MultiPage.Mapping.backward_highlight_CheckBox.Value Then
     While backward_input_link(0) <> ""
        backward_input_link = find_backward(backward_input_link, xmlDom)
     Wend
    End If
    If ConsoleForm.Console_MultiPage.Mapping.forward_highlight_CheckBox.Value Then
     While forward_input_link(0) <> ""
        forward_input_link = find_forward(forward_input_link, xmlDom)
     Wend
    End If

    Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Highlight")
End Sub

'find backward trnsf name,mark blue
Function find_backward(input_link() As String, xmlDom As MSXML2.DOMDocument)
        Dim output_link() As String
        
        ReDim output_link(100)
        i = 0
        
        Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/MAPPING/CONNECTOR")
        For j = 0 To UBound(input_link)
        If input_link(j) <> "" Then
            analysis_result_end_at = ActiveSheet.Range("A65535").End(xlUp).row
            For k = 10 To analysis_result_end_at
                trnsf_name = ActiveSheet.Range("A" & k).Value
                If InStr(trnsf_name, "(") = 0 Then
                    If input_link(j) = trnsf_name Then
                         ActiveSheet.Range("A" & k).Interior.ColorIndex = 42
                    End If
                Else
                    If input_link(j) = Mid(trnsf_name, 1, InStr(trnsf_name, "(") - 1) Then
                        ActiveSheet.Range("A" & k).Interior.ColorIndex = 42
                    End If
                End If
            Next
            For Each xmlNode In xmlNodeList
               If xmlNode.attributes.getNamedItem("TOINSTANCE").nodeValue = input_link(j) Then
                    For i_check = 0 To i - 1
                        If xmlNode.attributes.getNamedItem("FROMINSTANCE").nodeValue = output_link(i_check) Then
                            exist_flg = 1
                        End If
                    Next
                    If exist_flg = 0 Then
                        output_link(i) = xmlNode.attributes.getNamedItem("FROMINSTANCE").nodeValue
                        i = i + 1
                    End If
                    exist_flg = 0
                End If
            Next
        End If
        Next
        find_backward = output_link
End Function

'find forward trnsf name, mark pink
Function find_forward(input_link() As String, xmlDom As MSXML2.DOMDocument)
        Dim output_link() As String
        
        ReDim output_link(100)
        i = 0
        
        Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/MAPPING/CONNECTOR")
        For j = 0 To UBound(input_link)
        If input_link(j) <> "" Then
            analysis_result_end_at = ActiveSheet.Range("A65535").End(xlUp).row
            For k = 10 To analysis_result_end_at
                trnsf_name = ActiveSheet.Range("A" & k).Value
                If InStr(trnsf_name, "(") = 0 Then
                    If input_link(j) = trnsf_name Then
                         ActiveSheet.Range("A" & k).Interior.ColorIndex = 38
                    End If
                Else
                    If input_link(j) = Mid(trnsf_name, 1, InStr(trnsf_name, "(") - 1) Then

                        ActiveSheet.Range("A" & k).Interior.ColorIndex = 38
                    End If
                End If
            Next
            For Each xmlNode In xmlNodeList
               If xmlNode.attributes.getNamedItem("FROMINSTANCE").nodeValue = input_link(j) Then
                    For i_check = 0 To i - 1
                        If xmlNode.attributes.getNamedItem("TOINSTANCE").nodeValue = output_link(i_check) Then
                            exist_flg = 1
                        End If
                    Next
                    If exist_flg = 0 Then
                        output_link(i) = xmlNode.attributes.getNamedItem("TOINSTANCE").nodeValue
                        i = i + 1
                    End If
                    exist_flg = 0
                End If
            Next
        End If
        Next
        find_forward = output_link
End Function

'----------------------------------
'call by button click
'----------------------------------
'Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

Public Sub Sub_Prepare_Propagate(xmlDom As MSXML2.DOMDocument)
On Error GoTo FATAL_ERROR
    'Check mapping XML DOM is vaild
    If mapping_select_file_flg = 0 Or xmlDom Is Nothing Then
        Call Sub_OkOnly_Msgbox("Please click 'Select A File' first.")
        Exit Sub
    End If
    
    Dim rn As Range
    Dim urn As Range
    Dim trnsf_name() As String
    Dim trnsf_type() As String
    Dim port_name() As String
    Dim port_data_type() As String
    Dim port_prec() As String
    Dim port_scale() As String
       
    'count trnsf and port first
    t_c = 0
    p_c = 0
    
    For Each rn In Selection
        If rn.Column = "1" Then
            Select Case ActiveSheet.Cells(rn.row, 2).Value
            'check supported type
            Case "Sorter", "Expression", "Aggregator", "Joiner", "Source Qualifier", "Filter", "Router", "Custom Transformation"
                t_c = t_c + 1
            Case Else
                Call Sub_OkOnly_Msgbox("Do Not Support To Propagate To This Kind Of Transformation.")
                ActiveSheet.Cells(rn.row, 1).Interior.ColorIndex = 3
                Exit Sub
            End Select
        End If
        If rn.Column = "4" Then
            p_c = p_c + 1
        End If
    Next
    
    ReDim trnsf_name(t_c), trnsf_type(t_c), port_name(p_c), port_data_type(p_c), port_prec(p_c), port_scale(p_c)
    
    If t_c = 0 Then
        Call Sub_OkOnly_Msgbox("Please Select Transformations!")
        Exit Sub
    End If
    
    prefix_flg = 0
    suffix_flg = 0
    'add prefix or suffix option
    
    If UBound(src_keep_port_name) = 0 Then
        If MsgBox("Do You Want To Add Unitive Prefix Or Suffix When Propagate?", vbYesNo, "TabularInfa") = vbYes Then
            If MsgBox("Yes to add prefix, No to add suffix.", vbYesNo) = vbYes Then
                add_prefix = InputBox("Please input the prefix you want to add.")
                add_prefix = LTrim(RTrim(add_prefix))
                If add_prefix = "" Then
                    Exit Sub
                Else
                    prefix_flg = 1
                End If
            Else
                add_suffix = InputBox("Please input the suffix you want to add.", "TabularInfa")
                add_suffix = LTrim(RTrim(add_suffix))
                If add_suffix = "" Then
                    Exit Sub
                Else
                    suffix_flg = 1
                End If
            End If
        End If
    End If

    t_c = 0
    p_c = 0
    For Each rn In Selection
        Select Case rn.Column
        Case "1"
            trnsf_name(t_c) = rn.Value
            trnsf_type(t_c) = Cells(rn.row, rn.Column + 1).Value
            t_c = t_c + 1
        Case "4"
            If prefix_flg = 1 Then
                port_name(p_c) = add_prefix + rn.Value
                fix_type = "prefix"
                fix_value = add_prefix
            Else
                If suffix_flg = 1 Then
                    port_name(p_c) = rn.Value + add_suffix
                    fix_type = "suffix"
                    fix_value = add_suffix
                Else
                    port_name(p_c) = rn.Value
                    fix_type = ""
                    fix_value = ""
                End If
            End If
            port_data_type(p_c) = Cells(rn.row, rn.Column + 1).Value
            port_prec(p_c) = Cells(rn.row, rn.Column + 2).Value
            port_scale(p_c) = Cells(rn.row, rn.Column + 3).Value
            p_c = p_c + 1
        Case Else
        End Select
    Next

    'check selection result
    If p_c = 0 And UBound(src_keep_port_name) = 0 Then
        Call Sub_OkOnly_Msgbox("Please select ports for propagetion!")
    Else
        If UBound(src_keep_port_name) <> 0 Then
            port_name = src_keep_port_name
            port_data_type = src_keep_port_data_type
            port_prec = src_keep_port_prec
            port_scale = src_keep_port_scale
            ReDim src_keep_port_name(0)
            ReDim src_keep_port_data_type(0)
            ReDim src_keep_port_prec(0)
            ReDim src_keep_port_scale(0)
        End If
    End If

    Call Sub_Propagate_Port(fix_type, fix_value, xmlDom, selected_trnsf_name_hist, selected_trnsf_type, trnsf_name, trnsf_type, port_name, port_data_type, port_prec, port_scale)
    
    Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Prepare_Propagate")
End Sub
'----------------------------------
'test case for ubound&lbound
'----------------------------------
Sub TEST()
Dim a() As String
ReDim a(0)
MsgBox UBound(a)
MsgBox LBound(a)
b = Array("1")
MsgBox UBound(b)
MsgBox LBound(b)
End Sub
