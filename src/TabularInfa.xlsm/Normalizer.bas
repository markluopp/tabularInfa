Attribute VB_Name = "Normalizer"
'----------------------------------
'mluo@merkleinc.com
'Version:
'2015-4-22 call by 'edit_mapping'
'2015-4-24 added walk_srcfield function to handl src field group
'----------------------------------
Public normalizer_flg As Integer
'0---need normalizer before update
'1---rdy to update
'3---forbid upaete
Public Sub Sub_Edit_Nrm(xmlDom As MSXML2.DOMDocument, nrm_name As String)
On Error GoTo FATAL_ERROR
     Dim xmlNode As MSXML2.IXMLDOMNode
     Dim xmlNodeList As MSXML2.IXMLDOMNodeList
     Dim chlNode As MSXML2.IXMLDOMNode
     Dim chiNodeList As MSXML2.IXMLDOMNodeList
     Dim test_chlNode As MSXML2.IXMLDOMNode
     
     output_at_row_1 = 10
     output_at_row_2 = 10
     'detect max column number
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
     
     If InStr(nrm_name, "(") = 0 Then
        reuseable_flg = 0
        Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/MAPPING/TRANSFORMATION")
     Else
        reuseable_flg = 1
        nrm_name = Mid(nrm_name, InStr(nrm_name, "(") + 1, Len(nrm_name) - InStr(nrm_name, "(") - 1)
        Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/TRANSFORMATION")
     End If
     
     'Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/MAPPING/TRANSFORMATION")
     Set xmlNode = Nothing
     Set chlNodeList = Nothing
        For Each xmlNode In xmlNodeList
          If xmlNode.attributes.getNamedItem("NAME").nodeValue = nrm_name Then
              Set chlNodeList = xmlNode.childNodes
                 For Each chlNode In chlNodeList
                    If chlNode.nodeName = "TRANSFORMFIELD" Then
                        port_name = chlNode.attributes.getNamedItem("NAME").nodeValue
                        port_datatype = chlNode.attributes.getNamedItem("DATATYPE").nodeValue
                        port_pre = chlNode.attributes.getNamedItem("PRECISION").nodeValue
                        port_scale = chlNode.attributes.getNamedItem("SCALE").nodeValue
                        port_type = chlNode.attributes.getNamedItem("PORTTYPE").nodeValue
                            
                        ActiveSheet.Range("D" & output_at_row_1).Value = port_name
                        ActiveSheet.Range("E" & output_at_row_1).Value = port_datatype
                        ActiveSheet.Range("F" & output_at_row_1).Value = port_pre
                        ActiveSheet.Range("G" & output_at_row_1).Value = port_scale
                        ActiveSheet.Range("H" & output_at_row_1).Value = port_type
                        output_at_row_1 = output_at_row_1 + 1
                    End If
                    If chlNode.nodeName = "SOURCEFIELD" Then
                        column_name = chlNode.attributes.getNamedItem("NAME").nodeValue
                        lvl = chlNode.attributes.getNamedItem("LEVEL").nodeValue
                        occurs = chlNode.attributes.getNamedItem("OCCURS").nodeValue
                        If chlNode.attributes.getNamedItem("FIELDTYPE").nodeValue = "GRPITEM" Then
                            data_type = ""
                            prec = ""
                            scl = ""
                        Else
                            data_type = chlNode.attributes.getNamedItem("DATATYPE").nodeValue
                            prec = chlNode.attributes.getNamedItem("PRECISION").nodeValue
                            scl = chlNode.attributes.getNamedItem("SCALE").nodeValue
                        End If

                        ActiveSheet.Range("I" & output_at_row_2).Value = column_name
                        ActiveSheet.Range("J" & output_at_row_2).Value = lvl
                        ActiveSheet.Range("K" & output_at_row_2).Value = occurs
                        ActiveSheet.Range("L" & output_at_row_2).Value = data_type
                        ActiveSheet.Range("M" & output_at_row_2).Value = prec
                        ActiveSheet.Range("N" & output_at_row_2).Value = scl
                        output_at_row_2 = output_at_row_2 + 1
                        If Not chlNode.FirstChild Is Nothing Then
                            Set test_chlNode = chlNode.FirstChild
                            While (Not test_chlNode Is Nothing)
                                'MsgBox test_chlNode.nodeName
                                On Error GoTo next1
                                column_name = test_chlNode.attributes.getNamedItem("NAME").nodeValue
                                'MsgBox "handling" + column_name
                                lvl = test_chlNode.attributes.getNamedItem("LEVEL").nodeValue
                                occurs = test_chlNode.attributes.getNamedItem("OCCURS").nodeValue
                                If test_chlNode.attributes.getNamedItem("FIELDTYPE").nodeValue = "GRPITEM" Then
                                    data_type = ""
                                    prec = ""
                                    scl = ""
                                Else
                                    data_type = test_chlNode.attributes.getNamedItem("DATATYPE").nodeValue
                                    prec = test_chlNode.attributes.getNamedItem("PRECISION").nodeValue
                                    scl = test_chlNode.attributes.getNamedItem("SCALE").nodeValue
                                End If
        
                                ActiveSheet.Range("I" & output_at_row_2).Value = column_name
                                ActiveSheet.Range("J" & output_at_row_2).Value = lvl
                                ActiveSheet.Range("K" & output_at_row_2).Value = occurs
                                ActiveSheet.Range("L" & output_at_row_2).Value = data_type
                                ActiveSheet.Range("M" & output_at_row_2).Value = prec
                                ActiveSheet.Range("N" & output_at_row_2).Value = scl
                                output_at_row_2 = output_at_row_2 + 1
                                'MsgBox column_name + "done"
                                Set test_chlNode = walk_srcfield(test_chlNode)
                                'MsgBox column_name + "get next node success"
                            Wend
next1:
                        End If
                    End If
                Next
            End If
        Next
            
        If chlNodeList Is Nothing Then
            Call Sub_OkOnly_Msgbox("Can not find specified normalizer '" + nrm_name + "'.")
            Exit Sub
        End If
        
        ActiveSheet.Range("D" + CStr(ActiveSheet.Range("D65535").End(xlUp).row) + ":" + Chr(header_end_at + 64) + "9").Columns.AutoFit
        Set xmlNode = Nothing
        Set xmlNodeList = Nothing
        Set chlNode = Nothing
        Set chlNodeList = Nothing
        
        normalizer_flg = 0
        
        Call Sub_Hint_Box_Add(Format(Time, "hh:mm:ss") + ": You are editing " + nrm_name + " and two port layouts have displayed.You Can ONLY Modify The Second One As You Want, then click 'Generate Normalizer Ports Layout' to generate first one." + vbLf)
        Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
        Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Edit_Nrm")

End Sub
'----------------------------------
'Assitant function to traverse src fields for nrm node
'----------------------------------
Function walk_srcfield(node As MSXML2.IXMLDOMNode) As MSXML2.IXMLDOMNode

     If Not node.FirstChild Is Nothing Then
        Set walk_srcfield = node.FirstChild
     Else
        If Not node.nextSibling Is Nothing Then
            Set walk_srcfield = node.nextSibling
        Else
            If node.parentNode.nextSibling Is Nothing Then
                Set walk_srcfield = Nothing
            Else
                Set walk_srcfield = node.parentNode.nextSibling
            End If
        End If
     End If

End Function

'----------------------------------
'mluo@merkleinc.com
'Version:
'2015-4-23 intail version
'2014-4-24 divide this sub into 'normal' and 'update'
'----------------------------------
Public Sub Sub_Normal_Nrm()
On Error GoTo FATAL_ERROR
    'check normalizer tab
     output_at_row_2 = 10
     end_at_row_2 = ActiveSheet.Range("I65535").End(xlUp).row
     zero_lvl_flg = 0
     nonzero_lvl_flg = 0
     
     For check_at_row_2 = 10 To end_at_row_2
        'check column name
        If ActiveSheet.Range("I" & check_at_row_2).Value = "" Then
            Call Sub_OkOnly_Msgbox("Column name can not be null!")
            ActiveSheet.Range("I" & check_at_row_2).Interior.ColorIndex = 3
            Exit Sub
        End If
        'check data type
        If ActiveSheet.Range("L" & check_at_row_2).Value <> "number" And ActiveSheet.Range("L" & check_at_row_2).Value <> "string" And ActiveSheet.Range("L" & check_at_row_2).Value <> "nstring" And ActiveSheet.Range("L" & check_at_row_2).Value <> "" Then
            Call Sub_OkOnly_Msgbox("Invalid datatype. Only support number/string/nstring in NORMALIZER tab!")
            ActiveSheet.Range("L" & check_at_row_2).Interior.ColorIndex = 3
            Exit Sub
        End If
        'check occurs
        If ActiveSheet.Range("K" & check_at_row_2).Value < 0 Or (Not IsNumeric(ActiveSheet.Range("K" & check_at_row_2).Value)) Then
            Call Sub_OkOnly_Msgbox("Invalid Occurs!")
            ActiveSheet.Range("K" & check_at_row_2).Interior.ColorIndex = 3
            Exit Sub
        End If
        'hygeian scale
        If ActiveSheet.Range("L" & check_at_row_2).Value <> "number" Then
            ActiveSheet.Range("N" & check_at_row_2).Value = 0
        End If
        
        If ActiveSheet.Range("J" & check_at_row_2).Value = "" Then
            Call Sub_OkOnly_Msgbox("Level CAN NOT BE NULL!")
            ActiveSheet.Range("J" & check_at_row_2).Interior.ColorIndex = 3
            Exit Sub
        End If
        If ActiveSheet.Range("J" & check_at_row_2).Value = 0 Then
            zero_lvl_flg = 1
        End If
        If ActiveSheet.Range("J" & check_at_row_2).Value <> 0 Then
            nonzero_lvl_flg = 1
        End If
        'detect zero and nonzero both exist in level
        If zero_lvl_flg = 1 And nonzero_lvl_flg = 1 Then
            Call Sub_OkOnly_Msgbox("Invalid level!")
            ActiveSheet.Range("J" & check_at_row_2).Interior.ColorIndex = 3
            Exit Sub
        End If
     Next check_at_row_2
    

    
    'update port tab
    output_at_row_1 = 10
    end_at_row_1 = ActiveSheet.Range("D65535").End(xlUp).row
    If end_at_row_1 < 10 Then
        end_at_row_1 = 10
    End If
    'clean history
    ActiveSheet.Range("D10:H" & end_at_row_1).Clear
    If zero_lvl_flg = 1 And nonzero_lvl_flg = 0 Then
    'all zero level
        'INPUT port
        For check_at_row_2 = 10 To end_at_row_2
            If ActiveSheet.Range("K" & check_at_row_2).Value < 2 Then
                ActiveSheet.Range("D" & output_at_row_1) = ActiveSheet.Range("I" & check_at_row_2) + "_in"
                If ActiveSheet.Range("L" & output_at_row_2) = "number" Then
                    ActiveSheet.Range("E" & output_at_row_1) = "decimal"
                Else
                    ActiveSheet.Range("E" & output_at_row_1) = ActiveSheet.Range("L" & check_at_row_2)
                End If
                
                ActiveSheet.Range("F" & output_at_row_1) = ActiveSheet.Range("M" & check_at_row_2)
                ActiveSheet.Range("G" & output_at_row_1) = ActiveSheet.Range("N" & check_at_row_2)
                ActiveSheet.Range("H" & output_at_row_1) = "INPUT"
                output_at_row_1 = output_at_row_1 + 1
            Else
                For port_count = 1 To ActiveSheet.Range("K" & check_at_row_2).Value
                    ActiveSheet.Range("D" & output_at_row_1) = ActiveSheet.Range("I" & check_at_row_2) + "_in" + CStr(port_count)
                    If ActiveSheet.Range("L" & output_at_row_2) = "number" Then
                        ActiveSheet.Range("E" & output_at_row_1) = "decimal"
                    Else
                        ActiveSheet.Range("E" & output_at_row_1) = ActiveSheet.Range("L" & check_at_row_2)
                    End If
                    
                    ActiveSheet.Range("F" & output_at_row_1) = ActiveSheet.Range("M" & check_at_row_2)
                    ActiveSheet.Range("G" & output_at_row_1) = ActiveSheet.Range("N" & check_at_row_2)
                    ActiveSheet.Range("H" & output_at_row_1) = "INPUT"
                    output_at_row_1 = output_at_row_1 + 1
                Next
            End If
        Next check_at_row_2
        'OUTPUT port
        For check_at_row_2 = 10 To end_at_row_2
            ActiveSheet.Range("D" & output_at_row_1) = ActiveSheet.Range("I" & check_at_row_2)
            If ActiveSheet.Range("L" & output_at_row_2) = "number" Then
                ActiveSheet.Range("E" & output_at_row_1) = "decimal"
            Else
                ActiveSheet.Range("E" & output_at_row_1) = ActiveSheet.Range("L" & check_at_row_2)
            End If
                
            ActiveSheet.Range("F" & output_at_row_1) = ActiveSheet.Range("M" & check_at_row_2)
            ActiveSheet.Range("G" & output_at_row_1) = ActiveSheet.Range("N" & check_at_row_2)
            ActiveSheet.Range("H" & output_at_row_1) = "OUTPUT"
            output_at_row_1 = output_at_row_1 + 1
        Next check_at_row_2
        'GK
        For check_at_row_2 = 10 To end_at_row_2
            If ActiveSheet.Range("K" & check_at_row_2).Value > 1 Then
                ActiveSheet.Range("D" & output_at_row_1) = "GK_" + ActiveSheet.Range("I" & check_at_row_2)
                ActiveSheet.Range("E" & output_at_row_1) = "bigint"
                ActiveSheet.Range("F" & output_at_row_1) = "19"
                ActiveSheet.Range("G" & output_at_row_1) = "0"
                ActiveSheet.Range("H" & output_at_row_1) = "GENERATED KEY/OUTPUT"
                output_at_row_1 = output_at_row_1 + 1
                GoTo GCID_OUTPUT
            End If
        Next check_at_row_2
        'GCID
GCID_OUTPUT:
        For check_at_row_2 = 10 To end_at_row_2
            If ActiveSheet.Range("K" & check_at_row_2).Value > 1 Then
                ActiveSheet.Range("D" & output_at_row_1) = "GCID_" + ActiveSheet.Range("I" & check_at_row_2)
                ActiveSheet.Range("E" & output_at_row_1) = "integer"
                ActiveSheet.Range("F" & output_at_row_1) = "10"
                ActiveSheet.Range("G" & output_at_row_1) = "0"
                ActiveSheet.Range("H" & output_at_row_1) = "GENERATED COLUMN ID/OUTPUT"
                output_at_row_1 = output_at_row_1 + 1
            End If
        Next check_at_row_2
    Else
    'all nonzero level
    'INPUT PORT
    'OUTPUT PORT
    'GK
    'GCID
        Call Sub_OkOnly_Msgbox("Column Groups are not supported in this Tool. Please configure it in Desinger.")
        normalizer_flg = 3
        Exit Sub
    End If
    'Ready for update
    normalizer_flg = 1
    
        Call Sub_Hint_Box_Add(Format(Time, "hh:mm:ss") + ": First part has been generate by what you set in second part. Please verify if it's what you expect,then click 'Update This Transformation' to save changes." + vbLf)
        Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
        Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Normal_Nrm")
End Sub


Public Sub Sub_Update_Nrm(xmlDom As MSXML2.DOMDocument, nrm_name As String)
On Error GoTo FATAL_ERROR
     Dim newNode As MSXML2.IXMLDOMNode
     Dim xmlNodeList As MSXML2.IXMLDOMNodeList
     Dim chlNode As MSXML2.IXMLDOMNode
     Dim test_chlNode As MSXML2.IXMLDOMNode

     If normalizer_flg <> 1 Then
         Call Sub_OkOnly_Msgbox("Please click Normalizer before Update.")
         Exit Sub
     End If

'    MsgBox "ready to update"
'    Exit Sub
    
    end_at_row_1 = ActiveSheet.Range("D65535").End(xlUp).row
    end_at_row_2 = ActiveSheet.Range("I65535").End(xlUp).row
    column_offset = 0
    
    If reuseable_flg = 0 Then
        Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/MAPPING/TRANSFORMATION")
    Else
        Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/TRANSFORMATION")
    End If
    'Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/MAPPING/TRANSFORMATION")
    Set xmlNode = Nothing
    For Each xmlNode In xmlNodeList
     
              If xmlNode.attributes.getNamedItem("NAME").nodeValue = nrm_name Then
                
                    Set test_chlNode = xmlNode.selectSingleNode("SOURCEFIELD/SOURCEFIELD")
                    
                    If Not test_chlNode Is Nothing Then
                        Set test_chlNode = xmlNode.FirstChild
                        While test_chlNode.nodeName = "SOURCEFIELD"
                            Set test_chlNode = test_chlNode.nextSibling
                            test_chlNode.previousSibling.parentNode.removeChild test_chlNode.previousSibling
                        Wend
                    End If
                    
                    Set chlNode = xmlNode.FirstChild
                    'normalizer tab
                    For output_at_row_2 = 10 To end_at_row_2
                        
                        column_name = ActiveSheet.Range("I" & output_at_row_2).Value
                        lvl = CStr(ActiveSheet.Range("J" & output_at_row_2).Value)
                        occurs = CStr(ActiveSheet.Range("K" & output_at_row_2).Value)
                        data_type = ActiveSheet.Range("L" & output_at_row_2).Value
                        prec = CStr(ActiveSheet.Range("M" & output_at_row_2).Value)
                        scl = CStr(ActiveSheet.Range("N" & output_at_row_2).Value)
                        
                        
                        'Add column
                        If chlNode.nodeName <> "SOURCEFIELD" Then
                            Set newNode = xmlDom.createElement("SOURCEFIELD")
                                    
                                    Set src_attr = xmlDom.createAttribute("BUSINESSNAME")
                                    src_attr.Value = ""
                                    newNode.attributes.setNamedItem (src_attr)

                                    Set src_attr = xmlDom.createAttribute("DATATYPE")
                                    src_attr.Value = data_type
                                    newNode.attributes.setNamedItem (src_attr)

                                    Set src_attr = xmlDom.createAttribute("DESCRIPTION")
                                    src_attr.Value = ""
                                    newNode.attributes.setNamedItem (src_attr)

                                    Set src_attr = xmlDom.createAttribute("FIELDNUMBER")
                                    src_attr.Value = CStr(output_at_row_2 - 9)
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
                                    src_attr.Value = "NOT A KEY"
                                    newNode.attributes.setNamedItem (src_attr)

                                    Set src_attr = xmlDom.createAttribute("LENGTH")
                                    src_attr.Value = prec
                                    newNode.attributes.setNamedItem (src_attr)

                                    Set src_attr = xmlDom.createAttribute("LEVEL")
                                    src_attr.Value = lvl
                                    newNode.attributes.setNamedItem (src_attr)
                                    
                                    Set src_attr = xmlDom.createAttribute("NAME")
                                    src_attr.Value = column_name
                                    newNode.attributes.setNamedItem (src_attr)
                                    
                                    Set src_attr = xmlDom.createAttribute("NULLABLE")
                                    src_attr.Value = "NULL"
                                    newNode.attributes.setNamedItem (src_attr)
                                    
                                    Set src_attr = xmlDom.createAttribute("OCCURS")
                                    src_attr.Value = occurs
                                    newNode.attributes.setNamedItem (src_attr)
                                    
                                    Set src_attr = xmlDom.createAttribute("OFFSET")
                                    src_attr.Value = column_offset
                                    newNode.attributes.setNamedItem (src_attr)
                                    
                                    Set src_attr = xmlDom.createAttribute("PHYSICALLENGTH")
                                    src_attr.Value = prec
                                    newNode.attributes.setNamedItem (src_attr)
                                    
                                    Set src_attr = xmlDom.createAttribute("PHYSICALOFFSET")
                                    src_attr.Value = column_offset
                                    newNode.attributes.setNamedItem (src_attr)
                                    
                                    Set src_attr = xmlDom.createAttribute("PICTURETEXT")
                                    Select Case data_type
                                    Case "number"
                                        src_attr.Value = "9(" + prec + ")"
                                    Case "string"
                                        src_attr.Value = "X(" + prec + ")"
                                    Case "nstring"
                                        src_attr.Value = "N(" + prec + ")"
                                    Case Else
                                        Call Sub_OkOnly_Msgbox("Invalid data type in normalizer!")
                                        ActiveSheet.Range("L" & output_at_row_2).Interior.ColorIndex = 3
                                        Exit Sub
                                    newNode.attributes.setNamedItem (src_attr)
                                    End Select
                                    
                                    Set src_attr = xmlDom.createAttribute("PRECISION")
                                    src_attr.Value = prec
                                    newNode.attributes.setNamedItem (src_attr)
                                    
                                    Set src_attr = xmlDom.createAttribute("SCALE")
                                    src_attr.Value = scl
                                    newNode.attributes.setNamedItem (src_attr)
                                    
                                    Set src_attr = xmlDom.createAttribute("USAGE_FLAGS")
                                    src_attr.Value = ""
                                    newNode.attributes.setNamedItem (src_attr)
                            
                            chlNode.parentNode.insertBefore newNode, chlNode
                            Set chlNode = chlNode.previousSibling
                        
                        Else
                            Select Case data_type
                                Case "number"
                                        pic_text = "9(" + prec + ")"
                                Case "string"
                                        pic_text = "X(" + prec + ")"
                                Case "nstring"
                                        pic_text = "N(" + prec + ")"
                                Case Else
                                        Call Sub_OkOnly_Msgbox("Invalid data type in normalizer!")
                                        ActiveSheet.Range("L" & output_at_row_2).Interior.ColorIndex = 3
                                    Exit Sub
                            End Select
                            chlNode.attributes.getNamedItem("NAME").nodeValue = column_name
                            chlNode.attributes.getNamedItem("DATATYPE").nodeValue = data_type
                            chlNode.attributes.getNamedItem("PRECISION").nodeValue = prec
                            chlNode.attributes.getNamedItem("LENGTH").nodeValue = prec
                            chlNode.attributes.getNamedItem("PHYSICALLENGTH").nodeValue = prec
                            chlNode.attributes.getNamedItem("SCALE").nodeValue = scl
                            chlNode.attributes.getNamedItem("OFFSET").nodeValue = column_offset
                            chlNode.attributes.getNamedItem("PHYSICALOFFSET").nodeValue = column_offset
                            chlNode.attributes.getNamedItem("PICTURETEXT").nodeValue = pic_text
                        End If

                        column_offset = column_offset + prec
                        Set chlNode = chlNode.nextSibling
                            
                    Next output_at_row_2
                    
                    'Remove column
                    While chlNode.nodeName = "SOURCEFIELD"
                        Set chlNode = chlNode.nextSibling
                        chlNode.previousSibling.parentNode.removeChild chlNode.previousSibling
                    Wend
                    

            'port tab
            For output_at_row_1 = 10 To end_at_row_1
                        
                        port_name = ActiveSheet.Range("D" & output_at_row_1).Value
                        data_type = ActiveSheet.Range("E" & output_at_row_1).Value
                        prec = CStr(ActiveSheet.Range("F" & output_at_row_1).Value)
                        scl = CStr(ActiveSheet.Range("G" & output_at_row_1).Value)
                        port_type = ActiveSheet.Range("H" & output_at_row_1).Value
                        
                        
                        'Add column
                        If chlNode.nodeName <> "TRANSFORMFIELD" Then
                            Set newNode = xmlDom.createElement("TRANSFORMFIELD")
                                    
                                    Set trnsf_attr = xmlDom.createAttribute("DATATYPE")
                                    trnsf_attr.Value = data_type
                                    newNode.attributes.setNamedItem (trnsf_attr)

                                    Set trnsf_attr = xmlDom.createAttribute("DEFAULTVALUE")
                                    trnsf_attr.Value = ""
                                    newNode.attributes.setNamedItem (trnsf_attr)

                                    Set trnsf_attr = xmlDom.createAttribute("DESCRIPTION")
                                    trnsf_attr.Value = ""
                                    newNode.attributes.setNamedItem (trnsf_attr)

                                    Set trnsf_attr = xmlDom.createAttribute("NAME")
                                    trnsf_attr.Value = port_name
                                    newNode.attributes.setNamedItem (trnsf_attr)

                                    Set trnsf_attr = xmlDom.createAttribute("PICTURETEXT")
                                    trnsf_attr.Value = ""
                                    newNode.attributes.setNamedItem (trnsf_attr)

                                    Set trnsf_attr = xmlDom.createAttribute("PORTTYPE")
                                    trnsf_attr.Value = port_type
                                    newNode.attributes.setNamedItem (trnsf_attr)

                                    Set trnsf_attr = xmlDom.createAttribute("PRECISION")
                                    trnsf_attr.Value = prec
                                    newNode.attributes.setNamedItem (trnsf_attr)

                                    Set trnsf_attr = xmlDom.createAttribute("REF_SOURCE_FIELD")
                                    If InStr(port_name, "_in") = 0 Then
                                        If Mid(port_name, 1, 3) = "GK_" Then
                                            trnsf_attr.Value = Mid(port_name, 4, Len(port_name) - 3)
                                        Else
                                            If Mid(port_name, 1, 5) = "GCID_" Then
                                                trnsf_attr.Value = Mid(port_name, 6, Len(port_name) - 5)
                                            Else
                                                trnsf_attr.Value = port_name
                                            End If
                                        End If
                                    Else
                                        trnsf_attr.Value = Mid(port_name, 1, InStr(port_name, "_in") - 1)
                                    End If
                                    newNode.attributes.setNamedItem (trnsf_attr)

                                    Set trnsf_attr = xmlDom.createAttribute("SCALE")
                                    trnsf_attr.Value = scl
                                    newNode.attributes.setNamedItem (trnsf_attr)
                            
                            chlNode.parentNode.insertBefore newNode, chlNode
                            Set chlNode = chlNode.previousSibling
                        
                        Else
                            chlNode.attributes.getNamedItem("NAME").nodeValue = port_name
                            chlNode.attributes.getNamedItem("DATATYPE").nodeValue = data_type
                            chlNode.attributes.getNamedItem("PORTTYPE").nodeValue = port_type
                            chlNode.attributes.getNamedItem("PRECISION").nodeValue = prec
                            chlNode.attributes.getNamedItem("SCALE").nodeValue = scl
                            chlNode.attributes.getNamedItem("DEFAULTVALUE").nodeValue = ""
                            If InStr(port_name, "_in") = 0 Then
                                If Mid(port_name, 1, 3) = "GK_" Then
                                    chlNode.attributes.getNamedItem("REF_SOURCE_FIELD").nodeValue = Mid(port_name, 4, Len(port_name) - 3)
                                Else
                                    If Mid(port_name, 1, 5) = "GCID_" Then
                                        chlNode.attributes.getNamedItem("REF_SOURCE_FIELD").nodeValue = Mid(port_name, 6, Len(port_name) - 5)
                                    Else
                                        chlNode.attributes.getNamedItem("REF_SOURCE_FIELD").nodeValue = port_name
                                    End If
                                End If
                            Else
                                chlNode.attributes.getNamedItem("REF_SOURCE_FIELD").nodeValue = Mid(port_name, 1, InStr(port_name, "_in") - 1)
                            End If
                        End If
                            
                        Set chlNode = chlNode.nextSibling
                            
                    Next output_at_row_1
                    
                    'Remove column
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
            normalizer_flg = 0
            
            Call Sub_OkOnly_Msgbox("Complete update.")
            Call Sub_Hint_Box_Add(Format(Time, "hh:mm:ss") + ": Port changes for " + nrm_name + " have been updated to the XML file." + vbLf)
            Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
        Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Update_Nrm")
End Sub


