Sub Adjust_BOM_From_Baseline()

    Improve_Execution.ScreenUpdating
    
    For Row = 17 To 266
        If Cells(Row, 10).Value = "" And Cells(Row, 14).Value <> 0 Then
            'Copy from range(Cells(row + 1, 2), Cells(266, 10))
            'Paste formulas and formats to Cells(row, 2)
            Range(Cells(Row + 1, 2), Cells(266, 10)).Copy
            Cells(Row, 2).PasteSpecial Paste:=xlPasteFormulas
            Range(Cells(Row + 1, 2), Cells(266, 10)).Copy
            Cells(Row, 2).PasteSpecial Paste:=xlPasteFormats
            Row = Row - 1
        End If
    Next Row

    With Range("B17:B266")
        .Interior.Color = RGB(255, 255, 153)
        .Font.Bold = False
        .HorizontalAlignment = xlLeft
    End With
    For Row = 17 To 266
        If Cells(Row, 3).Interior.Color = RGB(255, 255, 0) Then
            Cells(Row, 2).Interior.Color = RGB(255, 255, 0)
        End If 
    Next Row

    With Range("J17:J266")
        .Interior.Color = RGB(255, 255, 153)
        .Font.Bold = False
        .HorizontalAlignment = xlRight
    End WIth

    Range("J17:J266").Copy
    Range("G17:G266").PasteSpecial xlPasteFormulas
    Range("J17:J266").Copy
    Range("G17:G266").PasteSpecial xlPasteFormats

    For Row = 266 To 17 Step -1
        If Cells(Row, 2).Value <> "" Then
            Range(Cells(Row + 2, 2), Cells(266, 2)).EntireRow.Hidden = True
            Exit For
        End If
    Exit For


    Range("B17:B266").Font.Strikethrough = False
    Range("B17:B266").IndentLevel = 0   
    Range("B17:B266").HorizontalAlignment = xlLeft
    Range("B17:B266").Borders(xlEdgeLeft).LineStyle = xlContinuous
    Range("B17:B266").Borders(xlEdgeLeft).Weight = xlThin

    Improve_Execution.Restore

End Sub