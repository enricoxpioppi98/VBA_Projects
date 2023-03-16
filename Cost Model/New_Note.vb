Sub New_Note()

    Improve_Execution.ScreenUpdating_And_Calculation
    
    If ActiveSheet.Name <> "Program Summary" Then
        MsgBox "Please, only run this macro in the Program Summary tab."
        Improve_Execution.Restore
        Exit Sub
    End If

    pw = "GCM2016SC"
    If ActiveSheet.ProtectContents = True Then
        Was_Protected = True
        ActiveSheet.Unprotect pw
    Else
        Was_Protected = False
    End If
    
    Range("98:98").EntireRow.Insert Shift:=xlDown
    Range("98:98").EntireRow.Insert Shift:=xlDown

    Rows_In_Notes_Section = Range("A85:V95").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    If Rows_In_Notes_Section > 85 Then
        Rows_In_Notes_Section = Rows_In_Notes_Section - 85
    Else
        Rows_In_Notes_Section = 1
    End If
    
    Range("85:" & (85 + Rows_In_Notes_Section)).Copy
    Range("98:98").Insert
    
    Range("95:95").Copy
    Range((98 + Rows_In_Notes_Section + 1) & ":" & (98 + Rows_In_Notes_Section + 1)).PasteSpecial xlPasteFormats
    Range((98 + Rows_In_Notes_Section + 1) & ":" & (98 + Rows_In_Notes_Section + 1)).Borders(xlEdgeTop).LineStyle = xlNone

    Worksheets("Input").Range("C5:E5").Copy
    Worksheets("Program Summary").Range("C85").PasteSpecial xlPasteAllExceptBorders

    If Flag_Protected = 1 Then
        ActiveSheet.Protect pw
    End If
    
    Improve_Execution.Restore

End Sub