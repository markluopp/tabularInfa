Attribute VB_Name = "Link"
'----------------------------------
'Purpose:Contain All Functions Use For Link Edit
'Version:
'2015-4-15 intail version
'2015-4-16 complete autolink_edit
'2015-4-17 complete autolink
'2015-4-18 complete update autolink
'2015-4-21 fix bug:#can not locate output node correctly#
'2015-4-21 fix bug:#can find src/tgt if it's an instance#
'2015-4-21 fix bug:#fail to count link rule#
'----------------------------------
Public Sub Sub_Edit_Link(xmlDom As MSXML2.DOMDocument)
On Error GoTo FATAL_ERROR

    'Check mapping XML DOM is vaild
    If mapping_select_file_flg = 0 Or xmlDom Is Nothing Then
        Call Sub_OkOnly_Msgbox("Please click 'Select A File' first.")
        Exit Sub
    End If
    
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim field_Node As MSXML2.IXMLDOMNode
    Dim xmlNodeList As MSXML2.IXMLDOMNodeList
    Dim field_NodeList As MSXML2.IXMLDOMNodeList
    
    'Assign XML filename
     fr_trnsf_name = ThisWorkbook.Sheets("autolink").Range("B5").Value
     to_trnsf_name = ThisWorkbook.Sheets("autolink").Range("B6").Value
     
     ActiveSheet.Range("A8").Value = fr_trnsf_name
     ActiveSheet.Range("D8").Value = to_trnsf_name
     
     output_at_row = 10
    'Clean history
     For i = ActiveSheet.UsedRange.Rows.Count To output_at_row Step -1
        ActiveSheet.Rows(i).Delete
     Next
     
     fr_flg = 0
     to_flg = 0
     
     'check if exist reuseable
     If InStr(fr_trnsf_name, "(") Then
        fr_trnsf_name_reuse = Mid(fr_trnsf_name, InStr(fr_trnsf_name, "(") + 1, Len(fr_trnsf_name) - InStr(fr_trnsf_name, "(") - 1)
        fr_trnsf_name = Mid(fr_trnsf_name, 1, InStr(fr_trnsf_name, "(") - 1)
    End If
     
     If InStr(to_trnsf_name, "(") Then
        to_trnsf_name_reuse = Mid(to_trnsf_name, InStr(to_trnsf_name, "(") + 1, Len(to_trnsf_name) - InStr(to_trnsf_name, "(") - 1)
        to_trnsf_name = Mid(to_trnsf_name, 1, InStr(to_trnsf_name, "(") - 1)
     End If
     
     'load exist columns
     Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/MAPPING/TRANSFORMATION")
     For Each xmlNode In xmlNodeList
        If xmlNode.attributes.getNamedItem("NAME").nodeValue = fr_trnsf_name Then
            ThisWorkbook.Sheets("autolink").Range("C5").Value = xmlNode.attributes.getNamedItem("TYPE").nodeValue
            output_at_row = 10
            'Normalizer has two kinds of field type(TRANSFORMFIELD/SOURCEFIELD)
            'Set field_NodeList = xmlNode.selectNodes("TRANSFORMFIELD")
            Set field_NodeList = xmlNode.childNodes
            For Each field_Node In field_NodeList
                If field_Node.nodeName = "TRANSFORMFIELD" Or field_Node.nodeName = "SOURCEFIELD" Then
                    ThisWorkbook.Sheets("autolink").Range("A" & output_at_row).Value = field_Node.attributes.getNamedItem("NAME").nodeValue
                    output_at_row = output_at_row + 1
                End If
            Next
        fr_flg = 1
        End If
        If xmlNode.attributes.getNamedItem("NAME").nodeValue = to_trnsf_name Then
            ThisWorkbook.Sheets("autolink").Range("C6").Value = xmlNode.attributes.getNamedItem("TYPE").nodeValue
            Set field_NodeList = xmlNode.selectNodes("TRANSFORMFIELD")
            output_at_row = 10
            For Each field_Node In field_NodeList
                ThisWorkbook.Sheets("autolink").Range("D" & output_at_row).Value = field_Node.attributes.getNamedItem("NAME").nodeValue
                output_at_row = output_at_row + 1
            Next
        to_flg = 1
        End If
     Next
     'if fr_trnsf/to_trnsf  reuse?
    If fr_flg <> 1 Or to_flg <> 1 Then
        Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/TRANSFORMATION")
        For Each xmlNode In xmlNodeList
            If fr_flg <> 1 Then
                If xmlNode.attributes.getNamedItem("NAME").nodeValue = fr_trnsf_name_reuse Then
                    ThisWorkbook.Sheets("autolink").Range("C5").Value = xmlNode.attributes.getNamedItem("TYPE").nodeValue
                    output_at_row = 10
                    'Normalizer has two kinds of field type(TRANSFORMFIELD/SOURCEFIELD)
                    'Set field_NodeList = xmlNode.selectNodes("TRANSFORMFIELD")
                    Set field_NodeList = xmlNode.childNodes
                    For Each field_Node In field_NodeList
                        If field_Node.nodeName = "TRANSFORMFIELD" Or field_Node.nodeName = "SOURCEFIELD" Then
                            ThisWorkbook.Sheets("autolink").Range("A" & output_at_row).Value = field_Node.attributes.getNamedItem("NAME").nodeValue
                            output_at_row = output_at_row + 1
                        End If
                    Next
                fr_flg = 1
                End If
            End If
            If to_flg <> 1 Then
                If xmlNode.attributes.getNamedItem("NAME").nodeValue = fr_trnsf_name_reuse Then
                    ThisWorkbook.Sheets("autolink").Range("C6").Value = xmlNode.attributes.getNamedItem("TYPE").nodeValue
                    Set field_NodeList = xmlNode.selectNodes("TRANSFORMFIELD")
                    output_at_row = 10
                    For Each field_Node In field_NodeList
                        ThisWorkbook.Sheets("autolink").Range("D" & output_at_row).Value = field_Node.attributes.getNamedItem("NAME").nodeValue
                        output_at_row = output_at_row + 1
                    Next
                to_flg = 1
                End If
            End If
         Next
    End If
    
    
    
    'if fr_trnsf is a src instance? to_trnsf is a tgt instance?
    If fr_flg <> 1 Or to_flg <> 1 Then
        Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/MAPPING/INSTANCE")
        For Each xmlNode In xmlNodeList
            If fr_flg <> 1 Then
                If xmlNode.attributes.getNamedItem("TRANSFORMATION_TYPE").nodeValue = "Source Definition" And fr_trnsf_name <> xmlNode.attributes.getNamedItem("TRANSFORMATION_NAME").nodeValue And xmlNode.attributes.getNamedItem("NAME").nodeValue = fr_trnsf_name Then
                    fr_trnsf_name_def = xmlNode.attributes.getNamedItem("TRANSFORMATION_NAME").nodeValue
                End If
            End If
            If to_flg <> 1 Then
                If xmlNode.attributes.getNamedItem("TRANSFORMATION_TYPE").nodeValue = "Target Definition" And to_trnsf_name <> xmlNode.attributes.getNamedItem("TRANSFORMATION_NAME").nodeValue And xmlNode.attributes.getNamedItem("NAME").nodeValue = to_trnsf_name Then
                    to_trnsf_name_def = xmlNode.attributes.getNamedItem("TRANSFORMATION_NAME").nodeValue
                End If
            End If
        Next
    End If
    
     'if src to trnsf?
     If fr_flg <> 1 Then
        Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/SOURCE")
        For Each xmlNode In xmlNodeList
        If xmlNode.attributes.getNamedItem("NAME").nodeValue = fr_trnsf_name Or xmlNode.attributes.getNamedItem("NAME").nodeValue = fr_trnsf_name_def Then
            ThisWorkbook.Sheets("autolink").Range("C5").Value = "Source Definition"
            Set field_NodeList = xmlNode.selectNodes("SOURCEFIELD")
            output_at_row = 10
            For Each field_Node In field_NodeList
                    ThisWorkbook.Sheets("autolink").Range("A" & output_at_row).Value = field_Node.attributes.getNamedItem("NAME").nodeValue
                    output_at_row = output_at_row + 1
            Next
        fr_flg = 1
        End If
        Next
    End If
    
     'if trnsf to tgt?
     If to_flg <> 1 Then
        Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/TARGET")
        For Each xmlNode In xmlNodeList
        If xmlNode.attributes.getNamedItem("NAME").nodeValue = to_trnsf_name Or xmlNode.attributes.getNamedItem("NAME").nodeValue = to_trnsf_name_def Then
            ThisWorkbook.Sheets("autolink").Range("C6").Value = "Target Definition"
            Set field_NodeList = xmlNode.selectNodes("TARGETFIELD")
            output_at_row = 10
            For Each field_Node In field_NodeList
                    ThisWorkbook.Sheets("autolink").Range("D" & output_at_row).Value = field_Node.attributes.getNamedItem("NAME").nodeValue
                    output_at_row = output_at_row + 1
            Next
        to_flg = 1
        End If
        Next
    End If
     
    If fr_flg <> 1 Then
        MsgBox "Can not find " + fr_trnsf_name + "!"
        Exit Sub
    End If
    
    If to_flg <> 1 Then
        MsgBox "Can not find " + to_trnsf_name + "!"
        Exit Sub
    End If
    
'     load existed connectors
'     fr_trnsf_name = ThisWorkbook.Sheets("autolink").Range("B5").Value
'     to_trnsf_name = ThisWorkbook.Sheets("autolink").Range("B6").Value
     
    output_at_row = 10
    Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/MAPPING/CONNECTOR")
    For Each xmlNode In xmlNodeList
        If xmlNode.attributes.getNamedItem("FROMINSTANCE").nodeValue = fr_trnsf_name And xmlNode.attributes.getNamedItem("TOINSTANCE").nodeValue = to_trnsf_name Then
            ThisWorkbook.Sheets("autolink").Range("B" & output_at_row).Value = xmlNode.attributes.getNamedItem("FROMFIELD").nodeValue
            ThisWorkbook.Sheets("autolink").Range("C" & output_at_row).Value = xmlNode.attributes.getNamedItem("TOFIELD").nodeValue
            output_at_row = output_at_row + 1
        End If
    Next
    
    Call Sub_Hint_Box_Set(Format(Time, "hh:mm:ss") + ": Jump to 'autolink' tab to edit links between " + fr_trnsf_name + " and " + to_trnsf_name + vbLf)
    Call Sub_Hint_Box_Add("You can set the 'Link Rule' at top right corner, then Click 'Try Simulative AutoLink' to generate links according to what you set." + vbLf)
    Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
    Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Edit_Link")
    
 End Sub


Public Sub Sub_Autolink()
On Error GoTo FATAL_ERROR
     'Dim fr_field_name() As String
     'Dim to_field_name() As String
     Dim fr_prefix() As String
     Dim fr_suffix() As String
     Dim to_prefix() As String
     Dim to_suffix() As String



'    MsgBox fr_field_name(0)
'    MsgBox to_field_name(2)

    'read link rules
    link_rule_1 = [iv6].End(xlToLeft).Column
    link_rule_2 = [iv7].End(xlToLeft).Column
    link_rule_3 = [iv8].End(xlToLeft).Column
    link_rule_4 = [iv9].End(xlToLeft).Column
    
'    MsgBox link_rule_1
'    MsgBox link_rule_2
'    MsgBox link_rule_3
'    MsgBox link_rule_4
'    Exit Sub
    'count link rules
    If link_rule_1 > link_rule_2 Then
        link_rule_12 = link_rule_1 - 7
    Else
        link_rule_12 = link_rule_2 - 7
    End If
    
    If link_rule_3 > link_rule_4 Then
        link_rule_34 = link_rule_3 - 7
    Else
        link_rule_34 = link_rule_4 - 7
    End If
    
    If link_rule_12 > link_rule_34 Then
        link_rule = link_rule_12
    Else
        link_rule = link_rule_34
    End If
    
        ReDim fr_prefix(link_rule)
        ReDim fr_suffix(link_rule)
        ReDim to_prefix(link_rule)
        ReDim to_suffix(link_rule)
        For i = 1 To link_rule
            fr_prefix(i - 1) = Cells(6, (i + 7)).Value
            fr_suffix(i - 1) = Cells(7, (i + 7)).Value
            to_prefix(i - 1) = Cells(8, (i + 7)).Value
            to_suffix(i - 1) = Cells(9, (i + 7)).Value
        Next

    'MsgBox to_prefix(0)
    output_at_row = ThisWorkbook.Sheets("autolink").Range("B65535").End(xlUp).row + 1
    'MsgBox link_rule
'    Exit Sub
    'autolink by link rules
    For i = 1 To link_rule
        'read every port start
        For j = 10 To ThisWorkbook.Sheets("autolink").Range("A65535").End(xlUp).row
        fr_field_name = ThisWorkbook.Sheets("autolink").Range("A" & j).Value
        'match start perfix and suffix
        If UCase(Left(fr_field_name, Len(fr_prefix(i - 1)))) = UCase(fr_prefix(i - 1)) And UCase(Right(fr_field_name, Len(fr_suffix(i - 1)))) = UCase(fr_suffix(i - 1)) Then
            'MsgBox ThisWorkbook.Sheets("autolink").Range("C" & J).Value
            'Exit Sub
            'find in port end
            For k = 10 To ThisWorkbook.Sheets("autolink").Range("D65535").End(xlUp).row
            to_field_name = ThisWorkbook.Sheets("autolink").Range("D" & k).Value
            If UCase(Left(to_field_name, Len(to_prefix(i - 1)))) = UCase(to_prefix(i - 1)) And UCase(Right(to_field_name, Len(to_suffix(i - 1)))) = UCase(to_suffix(i - 1)) Then
            'MsgBox Mid(fr_field_name, Len(fr_prefix(I - 1)) + 1, Len(fr_field_name) - Len(fr_prefix(I - 1)) - Len(fr_suffix(I - 1)))
            'Exit Sub
                If UCase(Mid(to_field_name, Len(to_prefix(i - 1)) + 1, Len(to_field_name) - Len(to_prefix(i - 1)) - Len(to_suffix(i - 1)))) = UCase(Mid(fr_field_name, Len(fr_prefix(i - 1)) + 1, Len(fr_field_name) - Len(fr_prefix(i - 1)) - Len(fr_suffix(i - 1)))) Then
                    'check if port end exist in link history, end only has one input link
                    to_field_name_exist_flg = 0
                    For l = 10 To ThisWorkbook.Sheets("autolink").Range("C65535").End(xlUp).row
                        If ThisWorkbook.Sheets("autolink").Range("C" & l).Value = to_field_name Then
                            to_field_name_exist_flg = 1
                        End If
                    Next
                    If to_field_name_exist_flg = 0 Then
                        ThisWorkbook.Sheets("autolink").Range("B" & output_at_row).Value = fr_field_name
                        ThisWorkbook.Sheets("autolink").Range("C" & output_at_row).Value = to_field_name
                        output_at_row = output_at_row + 1
                    End If
                End If
            End If
            Next
        End If
        Next
    Next
    
    'default link by name
    'check enable option
    If ThisWorkbook.Sheets("autolink").Range("G5").Value = "Y" Then
        For j = 10 To ThisWorkbook.Sheets("autolink").Range("A65535").End(xlUp).row
            fr_field_name = ThisWorkbook.Sheets("autolink").Range("A" & j).Value
            For k = 10 To ThisWorkbook.Sheets("autolink").Range("D65535").End(xlUp).row
            to_field_name = ThisWorkbook.Sheets("autolink").Range("D" & k).Value
            If UCase(fr_field_name) = UCase(to_field_name) Then
                'check if port end exist in link history
                 to_field_name_exist_flg = 0
                        For l = 10 To ThisWorkbook.Sheets("autolink").Range("C65535").End(xlUp).row
                            If ThisWorkbook.Sheets("autolink").Range("C" & l).Value = to_field_name Then
                                to_field_name_exist_flg = 1
                            End If
                        Next
                        If to_field_name_exist_flg = 0 Then
                            ThisWorkbook.Sheets("autolink").Range("B" & output_at_row).Value = fr_field_name
                            ThisWorkbook.Sheets("autolink").Range("C" & output_at_row).Value = to_field_name
                            output_at_row = output_at_row + 1
                        End If
            End If
            Next
        Next
    End If
    
    Call Sub_OkOnly_Msgbox("Simulated link results have been created under 'Link Result Map'.")
    Call Sub_Hint_Box_Add(Format(Time, "hh:mm:ss") + ": Simulated result has displayed. Click 'Update Link Changes' to save changes if you have verified it's good.")
    Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
    Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Autolink")
End Sub
'----------------------------------
'Purpose:
'Version:
'2015-6-30 fix bug of <ERPINFO> exist
'----------------------------------
Public Sub Sub_Update_Autolink(xmlDom As MSXML2.DOMDocument)
'On Error GoTo FATAL_ERROR
    'Check mapping XML DOM is vaild
    If mapping_select_file_flg = 0 Or xmlDom Is Nothing Then
        ThisWorkbook.Sheets("edit_mapping").Activate
        Call Sub_OkOnly_Msgbox("Please click 'Select A File' first.")
        Exit Sub
    End If
     
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim newNode As MSXML2.IXMLDOMNode
    Dim xmlNodeList As MSXML2.IXMLDOMNodeList
    Dim ERPINFO_Node As MSXML2.IXMLDOMNode

    'Assign XML filename
     xml_filepath = ThisWorkbook.Sheets("autolink").Range("A2").Value
     xml_filename = ThisWorkbook.Sheets("autolink").Range("B4").Value
     fr_trnsf_name = ThisWorkbook.Sheets("autolink").Range("B5").Value
     to_trnsf_name = ThisWorkbook.Sheets("autolink").Range("B6").Value
     
    'check if exist reuseable
     If InStr(fr_trnsf_name, "(") Then
        fr_trnsf_name = Mid(fr_trnsf_name, 1, InStr(fr_trnsf_name, "(") - 1)
    End If
     
     If InStr(to_trnsf_name, "(") Then
        to_trnsf_name = Mid(to_trnsf_name, 1, InStr(to_trnsf_name, "(") - 1)
     End If
     
     output_at_row = 10
     'remove all existed link
     Set xmlNodeList = xmlDom.selectNodes("//POWERMART/REPOSITORY/FOLDER/MAPPING/CONNECTOR")
     For Each xmlNode In xmlNodeList
         If xmlNode.attributes.getNamedItem("FROMINSTANCE").nodeValue = fr_trnsf_name And xmlNode.attributes.getNamedItem("TOINSTANCE").nodeValue = to_trnsf_name Then
            xmlNode.parentNode.removeChild xmlNode
         End If
     Next
     
     Set xmlNode = xmlDom.selectSingleNode("//POWERMART/REPOSITORY/FOLDER/MAPPING")
     'Set locate_Node = xmlDom.selectSingleNode("//POWERMART/REPOSITORY/FOLDER/MAPPING/TARGETLOADORDER")
     'Set locate_Node = xmlNodeList.Item(xmlNodeList.Length - 1)
     Set locate_Node = xmlNodeList.Item(0)
     loc_add_flg = 0
     If locate_Node Is Nothing Then
        Set locate_Node = xmlDom.createElement("CONNECTOR_LOC")
        xmlNode.appendChild locate_Node
        Set locate_Node = xmlDom.selectSingleNode("//POWERMART/REPOSITORY/FOLDER/MAPPING/CONNECTOR_LOC")
        If locate_Node Is Nothing Then
            MsgBox "Can not find output location for connector node!"
            Exit Sub
        Else
            loc_add_flg = 1
            'Check if <ERPINFO/> Exist
            Set ERPINFO_Node = Nothing
            Set ERPINFO_Node = xmlNode.selectSingleNode("ERPINFO")
            If Not ERPINFO_Node Is Nothing Then
            'MsgBox "1"
                xmlNode.removeChild ERPINFO_Node
                xmlNode.appendChild ERPINFO_Node
            End If
        End If
     End If
     'MsgBox xmlNodeList.Length
     'MsgBox locate_Node.nodeName
     'Exit Sub
     'walk link result
     For output_at_row = 10 To ThisWorkbook.Sheets("autolink").Range("B65535").End(xlUp).row
        'check if fr_field exist
        fr_field_flg = 0
        For H = 10 To ThisWorkbook.Sheets("autolink").Range("A65535").End(xlUp).row
            If ThisWorkbook.Sheets("autolink").Range("A" & H).Value = ThisWorkbook.Sheets("autolink").Range("B" & output_at_row).Value Then
                fr_field_flg = 1
            End If
        Next
        If fr_field_flg = 0 Then
            ThisWorkbook.Sheets("autolink").Cells(output_at_row, "B").Interior.ColorIndex = 3
            MsgBox "Can not find this column name!"
            Exit Sub
        End If
        'check if to_field exist
        to_field_flg = 0
        For j = 10 To ThisWorkbook.Sheets("autolink").Range("D65535").End(xlUp).row
            If ThisWorkbook.Sheets("autolink").Range("D" & j).Value = ThisWorkbook.Sheets("autolink").Range("C" & output_at_row).Value Then
                to_field_flg = 1
            End If
        Next
        If to_field_flg = 0 Then
            ThisWorkbook.Sheets("autolink").Cells(output_at_row, "C").Interior.ColorIndex = 3
            MsgBox "Can not find this column name!"
            Exit Sub
        End If
        
            'check if two link have one same end
            For l = 10 To ThisWorkbook.Sheets("autolink").Range("C65535").End(xlUp).row
                If ThisWorkbook.Sheets("autolink").Range("C" & l).Value = ThisWorkbook.Sheets("autolink").Range("C" & output_at_row).Value And l <> output_at_row Then
                    ThisWorkbook.Sheets("autolink").Cells(output_at_row, "C").Interior.ColorIndex = 3
                    ThisWorkbook.Sheets("autolink").Cells(l, "C").Interior.ColorIndex = 3
                    MsgBox "Two links have same end!"
                    Exit Sub
                End If
            Next
                    
            'create link node
            Set newNode = xmlDom.createElement("CONNECTOR")
            
            Set con_attr = xmlDom.createAttribute("FROMFIELD")
            con_attr.Value = ThisWorkbook.Sheets("autolink").Range("B" & output_at_row).Value
            newNode.attributes.setNamedItem (con_attr)
            
            Set con_attr = xmlDom.createAttribute("FROMINSTANCE")
            con_attr.Value = fr_trnsf_name
            newNode.attributes.setNamedItem (con_attr)
            
            Set con_attr = xmlDom.createAttribute("FROMINSTANCETYPE")
            con_attr.Value = ThisWorkbook.Sheets("autolink").Range("C5").Value
            newNode.attributes.setNamedItem (con_attr)
            
            Set con_attr = xmlDom.createAttribute("TOFIELD")
            con_attr.Value = ThisWorkbook.Sheets("autolink").Range("C" & output_at_row).Value
            newNode.attributes.setNamedItem (con_attr)
            
            Set con_attr = xmlDom.createAttribute("TOINSTANCE")
            con_attr.Value = to_trnsf_name
            newNode.attributes.setNamedItem (con_attr)
            
            Set con_attr = xmlDom.createAttribute("TOINSTANCETYPE")
            con_attr.Value = ThisWorkbook.Sheets("autolink").Range("C6").Value
            newNode.attributes.setNamedItem (con_attr)
            
            'Set textNode = xmlDom.createTextNode(vbLf)
            'MsgBox locate_Node.parentNode.nodeName
            'Exit Sub
            locate_Node.parentNode.insertBefore newNode, locate_Node.nextSibling
            'xmlNode.insertBefore textNode, locate_Node
     Next
     
     'remove output_loc node
     If loc_add_flg = 1 Then
        locate_Node.parentNode.removeChild locate_Node
     End If
     
     xmlDom.Save xml_filepath + "/" + xml_filename
     
     Set xmlNode = Nothing
     Set xmlNodeList = Nothing
     Set locate_Node = Nothing
     
     Call Sub_OkOnly_Msgbox("Complete update.")
    ThisWorkbook.Sheets("edit_mapping").Activate
    Call Sub_Hint_Box_Add(Format(Time, "hh:mm:ss") + ": Link changes have been updated to the XML file.")
    Call Sub_Hint_Box_Add("------------------------------------------------------" + vbLf)
    Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Update_Autolink")
End Sub
