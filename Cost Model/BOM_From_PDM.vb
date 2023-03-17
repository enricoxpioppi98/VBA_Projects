Sub BOM_From_PDM()

    Improve_Execution.ScreenUpdating_And_Calculation

    'In row 5
    'Iterate 7 times from column A to L
    'If cell value is "Pos. #" or "Waste Rate" or "Component Location" or "Rev" or "State" or "Usage"
    'Remove that column
    For i = 1 To 7
        For column = 1 To 7
            If Cells(5, Column).Value = "Pos. #" Or Cells(5, Column).Value = "Waste Rate" Or Cells(5, Column).Value = "Component Location" Or Cells(5, Column).Value = "Rev" Or Cells(5, Column).Value = "State" Or Cells(5, Column).Value = "Usage" Then
                Cells(5, Column1).EntireColumn.Delete
            End If
        Next Column
    Next i

    'Starting from Row 6 all the way to the last row
    'in columns A and B
    'align to the left
        Range("A7:B300").HorizontalAlignment = xlLeft    

    'Remove Tools
        For Row = 7 To 300
            If Left(Cells(Row, 1).Value, 1) = 3 Then
                Cells(Row, 1).EntireRow.Delete
            End If
        Next Row

    'Starting from Row 6 all the way to the last row
    'in column A
    'indentaton level = value in column A - 1
        For Row = 7 To 300
            Cells(Row, 1).IndentLevel = Cells(Row, 1).Value - 1
        Next Row
    
    Improve_Execution.Restore

End Sub