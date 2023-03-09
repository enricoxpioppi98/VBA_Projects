Public Function IsInArray(ToBeFound As Variant, arr As Variant, NumberOfElements As Variant) As Variant
    Dim i
    For i = 1 To NumberOfElements
        If arr(i, 1) = ToBeFound Then
            arr(i, 1) = 0
            IsInArray = i
            Exit Function
        End If
    Next i
End Function
Sub BaselineBOM()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    xCurrentRow = ActiveCell.Row
    xCurrentColumn = ActiveCell.Column
    
    Range("A3:F143361").Font.Color = RGB(0, 0, 0)
    Range("A3:F143361").Font.Name = "Arial"
    Range("A3:F143361").Font.Size = 12
    
    Range("A3:D3").Interior.Color = RGB(242, 242, 242)
    Range("A4:D4").Interior.Color = RGB(217, 217, 217)
    Range("E3:F4").Interior.Color = RGB(255, 255, 153)
    
    Range("A3:A4,C3:C4").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    Range("E3:F4").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    
    Range("A3:F4").Copy
    Range("A5:F143361").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    'Count parts
        QIS_LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        QIS_NumberOfParts = QIS_LastRow - 2
        
        CM_LastRow = Cells(Rows.Count, 3).End(xlUp).Row
        CM_NumberOfParts = CM_LastRow - 2
        
    'arrays
        Dim QIS() As Variant
        ReDim QIS(1 To QIS_NumberOfParts, 1 To 2)
        For Row = 3 To QIS_LastRow
            i = Row - 2
            QIS(i, 1) = Cells(Row, 1).Value
            QIS(i, 2) = Cells(Row, 2).Value
        Next Row
        
        Dim CM() As Variant
        ReDim CM(1 To CM_NumberOfParts, 1 To 2)
        For Row = 3 To CM_LastRow
            i = Row - 2
            CM(i, 1) = Cells(Row, 3).Value
            CM(i, 2) = Cells(Row, 4).Value
        Next Row
        
    'final
        For i = 1 To CM_NumberOfParts
            Row = i + 2
            Cells(Row, 5).Value = CM(i, 1)                                          'paste CM BOM
            QIS_Element = IsInArray(CM(i, 1), QIS, QIS_NumberOfParts)
            If QIS_Element <> "" Then                                               'If CMitem is in QIS
                Cells(Row, 6).Value = QIS(QIS_Element, 2)                           'Paste QIS quantity and leave QISitem as 0
                If QIS(QIS_Element, 2) <> CM(i, 2) Then                             'If quantity changed
                    With Cells(Row, 5)
                        .Interior.Color = RGB(255, 255, 0)
                        .Font.Color = RGB(255, 0, 0)
                        .Font.Bold = True
                        .HorizontalAlignment = xlRight
                        .Font.Strikethrough = False
                    End With
                    With Cells(Row, 6)
                        .Interior.Color = RGB(255, 255, 0)
                        .Font.Color = RGB(255, 0, 0)
                        .Font.Bold = True
                        .HorizontalAlignment = xlRight
                        .Font.Strikethrough = False
                        .NumberFormat = "0.0000"
                    End With
                Else                                                                 'If NO change
                    With Cells(Row, 5)
                        .Interior.Color = RGB(255, 255, 153)
                        .Font.Color = RGB(255, 0, 0)
                        .HorizontalAlignment = xlLeft
                        .Font.Bold = False
                        .Font.Strikethrough = False
                    End With
                    With Cells(Row, 6)
                        .Interior.Color = RGB(255, 255, 153)
                        .Font.Color = RGB(255, 0, 0)
                        .HorizontalAlignment = xlRight
                        .Font.Bold = False
                        .Font.Strikethrough = False
                        .NumberFormat = "0.0000"
                    End With
                End If
            Else                                                                    'If CMitem is NOT in QIS (=removed)
                Cells(Row, 6).Value = ""                                            'Leave quantity blank
                With Cells(Row, 5)                                                  'Highlight green + strikethrough
                    .Interior.Color = RGB(226, 239, 218)
                    .Font.Color = RGB(255, 0, 0)
                    .Font.Bold = True
                    .HorizontalAlignment = xlRight
                    .Font.Strikethrough = True
                End With
            End If
            Final_LastRow = Row
        Next i
        For i = 1 To QIS_NumberOfParts
            If QIS(i, 1) <> 0 Then                                                  'If QISitem is NOT 0 (= is NOT in CM)
                Final_LastRow = Final_LastRow + 1
                Cells(Final_LastRow, 5).Value = QIS(i, 1)                           'Paste QISitem
                Cells(Final_LastRow, 6).Value = QIS(i, 2)                           'Paste QISitem quantity
                With Cells(Final_LastRow, 5)
                    .Interior.Color = RGB(255, 255, 0)
                    .Font.Color = RGB(255, 0, 0)
                    .Font.Bold = True
                    .HorizontalAlignment = xlRight
                    .Font.Strikethrough = False
                End With
                With Cells(Final_LastRow, 6)
                    .Interior.Color = RGB(255, 255, 0)
                    .Font.Color = RGB(255, 0, 0)
                    .Font.Bold = True
                    .HorizontalAlignment = xlRight
                    .Font.Strikethrough = False
                    .NumberFormat = "0.0000"
                End With
            End If
        Next i
        
        Cells(xCurrentRow, xCurrentColumn).Select
        
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
        
End Sub
