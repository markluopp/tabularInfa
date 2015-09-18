Private Sub worksheet_activate()
    ConsoleForm.Console_MultiPage.Value = 4
End Sub

Private Sub worksheet_change(ByVal Target As Range)
    If Target.Address(0, 0) = "A1" Then
        Select Case Target.Value
        Case "STT Style"
            ThisWorkbook.Sheets("Layout Hygiene").Range("A2") = "Field Name"
            ThisWorkbook.Sheets("Layout Hygiene").Range("B2") = "Data Type"
            ThisWorkbook.Sheets("Layout Hygiene").Range("C2") = ""
            ThisWorkbook.Sheets("Layout Hygiene").Range("D2") = ""
            ThisWorkbook.Sheets("Layout Hygiene").Range("E2") = ""
            ThisWorkbook.Sheets("Layout Hygiene").Range("F2") = ""
            ThisWorkbook.Sheets("Layout Hygiene").Range("A2:F2").Interior.ColorIndex = 15
            ThisWorkbook.Sheets("Layout Hygiene").Range("A2:F2").Font.FontStyle = "bold"
            'thisworkbook.sheets("Layout Hygiene").Range("A2:F2").Columns.AutoFit
        Case "FLD Style"
            ThisWorkbook.Sheets("Layout Hygiene").Range("A2") = "Field_Name"
            ThisWorkbook.Sheets("Layout Hygiene").Range("B2") = "Field_Business_Name"
            ThisWorkbook.Sheets("Layout Hygiene").Range("C2") = "Field_Description"
            ThisWorkbook.Sheets("Layout Hygiene").Range("D2") = "Field_Data_Type"
            ThisWorkbook.Sheets("Layout Hygiene").Range("E2") = "Precision"
            ThisWorkbook.Sheets("Layout Hygiene").Range("F2") = "Scale"
            ThisWorkbook.Sheets("Layout Hygiene").Range("A2:F2").Interior.ColorIndex = 15
            ThisWorkbook.Sheets("Layout Hygiene").Range("A2:F2").Font.FontStyle = "bold"
            'thisworkbook.sheets("Layout Hygiene").Range("A2:F2").Columns.AutoFit
        Case "Column List To Locate"
            ThisWorkbook.Sheets("Layout Hygiene").Range("A2") = "Locate These Field Names"
            ThisWorkbook.Sheets("Layout Hygiene").Range("B2") = ""
            ThisWorkbook.Sheets("Layout Hygiene").Range("C2") = ""
            ThisWorkbook.Sheets("Layout Hygiene").Range("D2") = ""
            ThisWorkbook.Sheets("Layout Hygiene").Range("E2") = ""
            ThisWorkbook.Sheets("Layout Hygiene").Range("F2") = ""
            ThisWorkbook.Sheets("Layout Hygiene").Range("A2:F2").Interior.ColorIndex = 15
            ThisWorkbook.Sheets("Layout Hygiene").Range("A2:F2").Font.FontStyle = "bold"
            'thisworkbook.sheets("Layout Hygiene").Range("A2:F2").Columns.AutoFit
        End Select
    End If
    
    If Target.Address(0, 0) = "H1" Then
        Select Case Target.Value
        Case "Source/Target Style"
            ThisWorkbook.Sheets("Layout Hygiene").Range("H2") = "Column Name"
            ThisWorkbook.Sheets("Layout Hygiene").Range("I2") = "Data Type"
            ThisWorkbook.Sheets("Layout Hygiene").Range("J2") = "Precision"
            ThisWorkbook.Sheets("Layout Hygiene").Range("K2") = "Scale"
            ThisWorkbook.Sheets("Layout Hygiene").Range("L2") = "Not Null"
            ThisWorkbook.Sheets("Layout Hygiene").Range("M2") = "Key Type"
            ThisWorkbook.Sheets("Layout Hygiene").Range("N2") = "Business Name"
            ThisWorkbook.Sheets("Layout Hygiene").Range("O2") = "Description"
            ThisWorkbook.Sheets("Layout Hygiene").Range("H2:O2").Interior.ColorIndex = 4
            ThisWorkbook.Sheets("Layout Hygiene").Range("H2:O2").Font.FontStyle = "bold"
            'thisworkbook.sheets("Layout Hygiene").Range("A2:F2").Columns.AutoFit
        Case "Transformation Style"
            ThisWorkbook.Sheets("Layout Hygiene").Range("H2") = "Port Name"
            ThisWorkbook.Sheets("Layout Hygiene").Range("I2") = "Port Data Type"
            ThisWorkbook.Sheets("Layout Hygiene").Range("J2") = "Precision"
            ThisWorkbook.Sheets("Layout Hygiene").Range("K2") = "Scale"
            ThisWorkbook.Sheets("Layout Hygiene").Range("L2") = ""
            ThisWorkbook.Sheets("Layout Hygiene").Range("M2") = ""
            ThisWorkbook.Sheets("Layout Hygiene").Range("N2") = ""
            ThisWorkbook.Sheets("Layout Hygiene").Range("O2") = ""
            ThisWorkbook.Sheets("Layout Hygiene").Range("H2:O2").Interior.ColorIndex = 3
            ThisWorkbook.Sheets("Layout Hygiene").Range("H2:O2").Font.FontStyle = "bold"
        End Select
    End If
End Sub

