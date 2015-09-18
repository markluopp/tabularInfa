Attribute VB_Name = "Layout_Hygiene"
'----------------------------------
'2015-4-29 demo by May Wang
'2015-4-29 initial version(only support NULL precision and NULL scale)
'2015-4-29 add clear_history()
'          fix bug#nstring change to string, bigint change to int#
'2015-5-26 add burst cell procedure
'----------------------------------
Public Sub Sub_Burst_Cell()
On Error GoTo FATAL_ERROR
    If ActiveSheet.Range("A1") <> "Column List To Locate" Then
        Call Sub_OkOnly_Msgbox("Please switch to 'Column List To Locate' first!")
        Exit Sub
    Else
        end_at_row = ActiveSheet.Range("A65535").End(xlUp).row
        output_at_row = 3
        For i = 3 To end_at_row
        cell_str = ActiveSheet.Range("A" & i).Value
            If InStr(cell_str, Chr(10)) <> 0 Then
               'MsgBox i
               'count fields in one cell
               field_count = Len(cell_str) - Len(Replace(cell_str, Chr(10), "")) + 1
               'MsgBox row_count
               field_length = 0
               For j = 1 To field_count - 1
                   Field = Mid(cell_str, 1, InStr(1, cell_str, Chr(10)))
                   cell_str = Replace(cell_str, Field, "")
                   Field = Replace(Replace(Field, Chr(10), ""), ",", "")
                   ActiveSheet.Range("B" & output_at_row).Value = Field
                   output_at_row = output_at_row + 1
               Next
               cell_str = Replace(cell_str, Chr(10), "")
               ActiveSheet.Range("B" & output_at_row).Value = cell_str
               output_at_row = output_at_row + 1
            Else
               ActiveSheet.Range("B" & output_at_row).Value = cell_str
               output_at_row = output_at_row + 1
            End If
        Next
    End If
    
    'move from column A to column B
    ActiveSheet.Range("A3:A" & end_at_row).Clear
    ActiveSheet.Range("A3:A" & output_at_row).Value = ActiveSheet.Range("B3:B" & output_at_row).Value
    ActiveSheet.Range("B3:B" & output_at_row).Clear
    
    Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Burst_Cell")
End Sub


Public Sub Sub_Clear_History()
On Error GoTo FATAL_ERROR
    input_end_at_row = ActiveSheet.Range("A65535").End(xlUp).row
    output_end_at_row = ActiveSheet.Range("H65535").End(xlUp).row
    If input_end_at_row >= 3 Then
        ActiveSheet.Range("A3:F" & input_end_at_row).Clear
    End If
    If output_end_at_row >= 3 Then
        ActiveSheet.Range("H3:O" & output_end_at_row).Clear
    End If
    Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Clear_History")
End Sub
'----------------------------------
'2015-4-29 fix bug of losing CurrentlyProcessedFileName
'----------------------------------
Public Sub Sub_Goto_Edit_Src()
On Error GoTo FATAL_ERROR
    add_file_name_flg = 0
    end_at_row = ThisWorkbook.Sheets("Layout Hygiene").Range("H65535").End(xlUp).row
    clear_end_at_row = ThisWorkbook.Sheets("edit_src").Range("A65535").End(xlUp).row
    If ThisWorkbook.Sheets("edit_src").Range("A" & clear_end_at_row).Value = "CurrentlyProcessedFileName" Then
        add_file_name_flg = 1
    End If
    If clear_end_at_row >= 10 Then
        ThisWorkbook.Sheets("edit_src").Range("A10:H" & clear_end_at_row).Clear
    End If
    For i = 3 To end_at_row
        ThisWorkbook.Sheets("edit_src").Range("A" & (i + 7)).Value = ThisWorkbook.Sheets("Layout Hygiene").Range("H" & i).Value
        ThisWorkbook.Sheets("edit_src").Range("B" & (i + 7)).Value = ThisWorkbook.Sheets("Layout Hygiene").Range("I" & i).Value
        ThisWorkbook.Sheets("edit_src").Range("C" & (i + 7)).Value = ThisWorkbook.Sheets("Layout Hygiene").Range("J" & i).Value
        ThisWorkbook.Sheets("edit_src").Range("D" & (i + 7)).Value = ThisWorkbook.Sheets("Layout Hygiene").Range("K" & i).Value
        ThisWorkbook.Sheets("edit_src").Range("E" & (i + 7)).Value = ThisWorkbook.Sheets("Layout Hygiene").Range("L" & i).Value
        ThisWorkbook.Sheets("edit_src").Range("F" & (i + 7)).Value = ThisWorkbook.Sheets("Layout Hygiene").Range("M" & i).Value
        ThisWorkbook.Sheets("edit_src").Range("G" & (i + 7)).Value = ThisWorkbook.Sheets("Layout Hygiene").Range("N" & i).Value
        ThisWorkbook.Sheets("edit_src").Range("H" & (i + 7)).Value = ThisWorkbook.Sheets("Layout Hygiene").Range("O" & i).Value
    Next
    If add_file_name_flg = 1 Then
        ThisWorkbook.Sheets("edit_src").Range("A" & (i + 7)).Value = "CurrentlyProcessedFileName"
        ThisWorkbook.Sheets("edit_src").Range("B" & (i + 7)).Value = "string"
        ThisWorkbook.Sheets("edit_src").Range("C" & (i + 7)).Value = "256"
        ThisWorkbook.Sheets("edit_src").Range("D" & (i + 7)).Value = "0"
        ThisWorkbook.Sheets("edit_src").Range("E" & (i + 7)).Value = "NULL"
        ThisWorkbook.Sheets("edit_src").Range("F" & (i + 7)).Value = "NOT A KEY"
        ThisWorkbook.Sheets("edit_src").Range("G" & (i + 7)).Value = ""
        ThisWorkbook.Sheets("edit_src").Range("H" & (i + 7)).Value = ""
    End If
    ThisWorkbook.Sheets("edit_src").Activate
    ThisWorkbook.Sheets("edit_src").Range("A9:H" & (i + 7)).Columns.AutoFit
    ThisWorkbook.Sheets("edit_src").Range("A9:H" & (i + 7)).Rows.AutoFit
Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Goto_Edit_Src")
End Sub

Public Sub Sub_Goto_Edit_Tgt()
On Error GoTo FATAL_ERROR
    end_at_row = ThisWorkbook.Sheets("Layout Hygiene").Range("H65535").End(xlUp).row
    clear_end_at_row = ThisWorkbook.Sheets("edit_tgt").Range("A65535").End(xlUp).row
    If clear_end_at_row >= 10 Then
        ThisWorkbook.Sheets("edit_tgt").Range("A10:H" & clear_end_at_row).Clear
    End If
    For i = 3 To end_at_row
        ThisWorkbook.Sheets("edit_tgt").Range("A" & (i + 7)).Value = ThisWorkbook.Sheets("Layout Hygiene").Range("H" & i).Value
        ThisWorkbook.Sheets("edit_tgt").Range("B" & (i + 7)).Value = ThisWorkbook.Sheets("Layout Hygiene").Range("I" & i).Value
        ThisWorkbook.Sheets("edit_tgt").Range("C" & (i + 7)).Value = ThisWorkbook.Sheets("Layout Hygiene").Range("J" & i).Value
        ThisWorkbook.Sheets("edit_tgt").Range("D" & (i + 7)).Value = ThisWorkbook.Sheets("Layout Hygiene").Range("K" & i).Value
        ThisWorkbook.Sheets("edit_tgt").Range("E" & (i + 7)).Value = ThisWorkbook.Sheets("Layout Hygiene").Range("L" & i).Value
        ThisWorkbook.Sheets("edit_tgt").Range("F" & (i + 7)).Value = ThisWorkbook.Sheets("Layout Hygiene").Range("M" & i).Value
        ThisWorkbook.Sheets("edit_tgt").Range("G" & (i + 7)).Value = ThisWorkbook.Sheets("Layout Hygiene").Range("N" & i).Value
        ThisWorkbook.Sheets("edit_tgt").Range("H" & (i + 7)).Value = ThisWorkbook.Sheets("Layout Hygiene").Range("O" & i).Value
    Next
    
    ThisWorkbook.Sheets("edit_tgt").Activate
    ThisWorkbook.Sheets("edit_tgt").Range("A9:H" & (i + 7)).Columns.AutoFit
    ThisWorkbook.Sheets("edit_tgt").Range("A9:H" & (i + 7)).Rows.AutoFit
Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Goto_Edit_Tgt")
End Sub
Public Sub Sub_Layout_Hygiene()
On Error GoTo FATAL_ERROR
    Dim datatype_str As String
    'get the last row number of input
    end_at_row = ActiveSheet.Range("A65535").End(xlUp).row
    
    For i = 3 To end_at_row
        'required column name
        If ActiveSheet.Range("A" & i).Value = "" Then
            ActiveSheet.Cells(i, "A").Interior.ColorIndex = 3
            MsgBox "Cloumn name is required!"
            Exit Sub
        End If
        'fix value first
        'clean column name and description by experience
        If ActiveSheet.Range("A1").Value = "FLD Style" Then
            'ActiveSheet.Range("H" & i).Value = ActiveSheet.Range("A" & i).Value
            ActiveSheet.Range("H" & i).Value = Replace(Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(ActiveSheet.Range("A" & i).Value, ")", ""), "(", ""), "-", " "), ".", " "), ":", " "), "&", " "), ":", " "), "/", " "), "+", " ")), " ", "_")
            ActiveSheet.Range("N" & i).Value = ActiveSheet.Range("B" & i).Value
            'ActiveSheet.Range("O" & i).Value = ActiveSheet.Range("C" & i).Value
            ActiveSheet.Range("O" & i).Value = Replace(Replace(Replace(Replace(Trim(ActiveSheet.Range("C" & i).Value), "&", " "), " ", ""), ChrW(8226), "&quot;"), """", "&quot;")
            ActiveSheet.Range("L" & i).Value = "NULL"
            ActiveSheet.Range("M" & i).Value = "NOT A KEY"
            datatype_str = LCase(ActiveSheet.Range("D" & i))
            prec = LCase(ActiveSheet.Range("E" & i))
            scle = LCase(ActiveSheet.Range("F" & i))
        Else
            ActiveSheet.Range("H" & i).Value = Replace(Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(ActiveSheet.Range("A" & i).Value, ")", ""), "(", ""), "-", " "), ".", " "), ":", " "), "&", " "), ":", " "), "/", " "), "+", " ")), " ", "_")
            ActiveSheet.Range("N" & i).Value = ""
            ActiveSheet.Range("O" & i).Value = ""
            ActiveSheet.Range("L" & i).Value = "NULL"
            ActiveSheet.Range("M" & i).Value = "NOT A KEY"
            datatype_str = LCase(ActiveSheet.Range("B" & i))
        End If
        
        If ActiveSheet.Range("H1").Value = "Transformation Style" Then
            ActiveSheet.Range("L" & i).Value = ""
            ActiveSheet.Range("M" & i).Value = ""
            ActiveSheet.Range("N" & i).Value = ""
            ActiveSheet.Range("O" & i).Value = ""
        End If
        
        'check data type
        Select Case True
        Case (InStr(datatype_str, "int") <> 0 Or InStr(datatype_str, "ineteger") <> 0) And (InStr(datatype_str, "big int") = 0 And InStr(datatype_str, "bigint") = 0)
             ActiveSheet.Cells(i, "I") = "int"
             GoTo length_hygiene
        Case InStr(datatype_str, "big int") <> 0 Or InStr(datatype_str, "bigint") <> 0
            ActiveSheet.Cells(i, "I") = "bigint"
            ActiveSheet.Cells(i, "J") = 19
            ActiveSheet.Cells(i, "K") = 0
            GoTo skip_legnth
        Case (InStr(datatype_str, "char") <> 0 Or InStr(datatype_str, "string") <> 0 Or InStr(datatype_str, "text") <> 0 Or InStr(datatype_str, "unicode") <> 0 Or InStr(datatype_str, "xxxxx") <> 0) And (InStr(datatype_str, "nchar") = 0 And InStr(datatype_str, "nvarchar") = 0)
            ActiveSheet.Cells(i, "I") = "string"
            GoTo length_hygiene
        Case InStr(datatype_str, "nchar") <> 0 Or InStr(datatype_str, "nvarchar") <> 0
            ActiveSheet.Cells(i, "I") = "nstring"
            GoTo length_hygiene
        Case InStr(datatype_str, "date") <> 0 Or InStr(datatype_str, "time") <> 0 Or InStr(datatype_str, "yyyymmdd") <> 0
            If ActiveSheet.Range("H1").Value = "Transformation Style" Then
                ActiveSheet.Cells(i, "I") = "date/time"
            Else
                ActiveSheet.Cells(i, "I") = "datetime"
            End If
            ActiveSheet.Cells(i, "J") = "29"
            ActiveSheet.Cells(i, "K") = "9"
            GoTo skip_legnth
        Case InStr(datatype_str, "num") <> 0 Or InStr(datatype_str, "decimal") <> 0 Or InStr(datatype_str, "float") <> 0
            ActiveSheet.Cells(i, "I") = "number"
            GoTo length_hygiene
        Case Else
            If ActiveSheet.Range("A1").Value = "FLD Style" Then
                ActiveSheet.Cells(i, "D").Interior.ColorIndex = 3
            Else
                ActiveSheet.Cells(i, "B").Interior.ColorIndex = 3
            End If
            MsgBox "Unrecognized Data Type!"
            Exit Sub
        End Select
        'no default length, use assistant function to split  out precision and scale
length_hygiene:
       
        If prec <> "" Then
            ActiveSheet.Cells(i, "J") = prec
            If scle <> "" Then
            ActiveSheet.Cells(i, "K") = scle
            Else
            ActiveSheet.Cells(i, "K") = 0
            End If
        Else
            If Not containNum(datatype_str) Then
                ActiveSheet.Cells(i, "J") = 10
                ActiveSheet.Cells(i, "K") = 0
            Else
                If InStr(1, datatype_str, ",") Then
                    ActiveSheet.Cells(i, "J") = getNum1(datatype_str)
                    ActiveSheet.Cells(i, "K") = getNum2(datatype_str)
                Else
                    ActiveSheet.Cells(i, "J") = getNum1(datatype_str)
                    ActiveSheet.Cells(i, "K") = 0
                End If
            End If
        End If
        'fixed length
skip_legnth:
    
    Next i
Exit Sub
FATAL_ERROR:
    Call Sub_Error_Handle("Sub_Layout_Hygiene")
End Sub

'assitant function of layout_hygiene to verify if the string contains number
Function containNum(str As String)
    Dim flag As Boolean
    flag = False
    For i = 1 To Len(str)
        If IsNumeric(Mid(str, i, 1)) Then
            flag = True
        End If
    Next
    containNum = flag
End Function
'assitant function of layout_hygiene to get precison
Function getNum1(str As String)
Dim ans As String
Dim flag As Boolean
ans = ""
flag = True

If InStr(str, "(") <> 0 Then
    For i = InStr(str, "(") To Len(str)
        If flag And IsNumeric(Mid(str, i, 1)) Then
                ans = ans & Mid(str, i, 1)
        End If
        
        If Mid(str, i, 1) = "," Then
            flag = False
        End If
    Next
Else
    For i = 1 To Len(str)
        If flag And IsNumeric(Mid(str, i, 1)) Then
                ans = ans & Mid(str, i, 1)
        End If
        
        If Mid(str, i, 1) = "," Then
            flag = False
        End If
    Next
End If

getNum1 = ans
End Function
'assitant function of layout_hygiene to get scale
Function getNum2(str As String)
Dim ans As String
Dim flag As Boolean
flag = True

ans = ""
For i = InStr(str, ",") To Len(str)
    If flag And IsNumeric(Mid(str, i, 1)) Then
            ans = ans & Mid(str, i, 1)
    End If
    If Mid(str, i, 1) = ")" Then
    flag = False
    End If
Next

getNum2 = ans
End Function



