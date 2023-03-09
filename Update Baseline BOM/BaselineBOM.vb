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

    'Improve Execution ------------------------------------------------------------------------------------------------------------------------------
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual

        Current_Row = ActiveCell.Row
        Current_Column = ActiveCell.Column
    '------------------------------------------------------------------------------------------------------------------------------------------------

    'Initial Formatting -----------------------------------------------------------------------------------------------------------------------------
        'BOM ----------------------------------------------------------------------------------------------------------------------------------------
            With Range("A3:B4")
                'Borders ----------------------------------------------------------------------------------------------------------------------------
                    .Borders(xlDiagonalDown).LineStyle = xlNone
                    .Borders(xlDiagonalUp).LineStyle = xlNone
                    .Borders(xlEdgeTop).LineStyle = xlNone
                    .Borders(xlEdgeBottom).LineStyle = xlNone
                    .Borders(xlInsideVertical).LineStyle = xlNone
                    .Borders(xlInsideHorizontal).LineStyle = xlNone
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                '------------------------------------------------------------------------------------------------------------------------------------
                'Font -------------------------------------------------------------------------------------------------------------------------------
                    .Font.Name = "Arial"
                    .Font.Size = 12
                    .Font.Color = RGB(0, 0, 0)
                    .Font.Bold = False
                '------------------------------------------------------------------------------------------------------------------------------------
            End With
            Range("A3:B3").Interior.Color = RGB(242, 242, 242)
            Range("A4:B4").Interior.Color = RGB(217, 217, 217)
        '--------------------------------------------------------------------------------------------------------------------------------------------
        'QIS ----------------------------------------------------------------------------------------------------------------------------------------
            Range("A3:B4").Copy
            Range("A3:D512").PasteSpecial Paste:=xlPasteFormats
        '--------------------------------------------------------------------------------------------------------------------------------------------
        'Final BOM ----------------------------------------------------------------------------------------------------------------------------------
            With Range("E3:F512")
                .Interior.Color = RGB(255, 255, 153)
                'Borders ----------------------------------------------------------------------------------------------------------------------------
                    .Borders(xlDiagonalDown).LineStyle = xlNone
                    .Borders(xlDiagonalUp).LineStyle = xlNone
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlInsideVertical).LineStyle = xlContinuous
                    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                '------------------------------------------------------------------------------------------------------------------------------------
                'Font -------------------------------------------------------------------------------------------------------------------------------
                    .Font.Name = "Arial"
                    .Font.Size = 12
                    .Font.Color = RGB(255, 0, 0)
                    .Font.Bold = False
                '------------------------------------------------------------------------------------------------------------------------------------
                .HorizontalAlignment = xlLeft
            End With
            With Range("B3:B512, D3:D512, F3:F512")
                .NumberFormat = "0.0000"
                .HorizontalAlignment = xlRight
            End With
        '--------------------------------------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------------------------------------

    'Count Parts ------------------------------------------------------------------------------------------------------------------------------------
        QIS_LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        QIS_NumberOfParts = QIS_LastRow - 2
        
        CM_LastRow = Cells(Rows.Count, 3).End(xlUp).Row
        CM_NumberOfParts = CM_LastRow - 2
    '------------------------------------------------------------------------------------------------------------------------------------------------

    'Arrays -----------------------------------------------------------------------------------------------------------------------------------------
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
    '------------------------------------------------------------------------------------------------------------------------------------------------

    'Calculate Final BOM ----------------------------------------------------------------------------------------------------------------------------

        For i = 1 To CM_NumberOfParts
            Row = i + 2
            Cells(Row, 5).Value = CM(i, 1)                                          'paste CM BOM
            QIS_Element = IsInArray(CM(i, 1), QIS, QIS_NumberOfParts)
            If QIS_Element <> "" Then                                               'If CMitem is in QIS
                Cells(Row, 6).Value = QIS(QIS_Element, 2)                           'Paste QIS quantity and leave QISitem as 0
                If QIS(QIS_Element, 2) <> CM(i, 2) Then                             'If quantity changed
                    'Bold, Align Right, Highlight Yellow  -------------------------------------------------------------------------------------------
                        With Range("E" & Row & ":F" & Row)
                            .Font.Bold = True
                            .HorizontalAlignment = xlRight
                            .Interior.Color = RGB(255, 255, 0)
                        End With
                    '--------------------------------------------------------------------------------------------------------------------------------
                End If
            Else                                                                    'If CM item is NOT in QIS (=removed)
                Cells(Row, 6).Value = ""                                            'Leave quantity blank
                'Bold, Align Right, Highlight Green, Strikethrough  ---------------------------------------------------------------------------------
                    With Range("E" & Row & ":F" & Row)
                        .Font.Bold = True
                        .HorizontalAlignment = xlRight
                        .Interior.Color = RGB(226, 239, 218)
                        .Font.Strikethrough = True
                    End With
                '------------------------------------------------------------------------------------------------------------------------------------
            End If
            Final_LastRow = Row
        Next i

        For i = 1 To QIS_NumberOfParts
            If QIS(i, 1) <> 0 Then                                                  'If QISitem is NOT 0 (= is NOT in CM)
                Final_LastRow = Final_LastRow + 1

                Cells(Final_LastRow, 5).Value = QIS(i, 1)                           'Paste QISitem
                Cells(Final_LastRow, 6).Value = QIS(i, 2)                           'Paste QISitem quantity

                'Bold, Align Right, Highlight Yellow  -------------------------------------------------------------------------------------------
                    With Range("E" & Final_LastRow & ":F" & Final_LastRow)
                        .Font.Bold = True
                        .HorizontalAlignment = xlRight
                        .Interior.Color = RGB(255, 255, 0)
                    End With
                '--------------------------------------------------------------------------------------------------------------------------------
            End If
        Next i
    '------------------------------------------------------------------------------------------------------------------------------------------------

    'Restore Execution ------------------------------------------------------------------------------------------------------------------------------
        Cells(Current_Row, Current_Column).Select
        
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    '------------------------------------------------------------------------------------------------------------------------------------------------
    
End Sub
