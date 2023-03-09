Sub CleanFinal()    
    
    'Improve Execution ------------------------------------------------------------------------------------------------------------------------------
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual

        Current_Row = ActiveCell.Row
        Current_Column = ActiveCell.Column
    '------------------------------------------------------------------------------------------------------------------------------------------------

    Range("E3:F512").ClearContents

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

    'Restore Execution ------------------------------------------------------------------------------------------------------------------------------
        Cells(Current_Row, Current_Column).Select
        
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    '------------------------------------------------------------------------------------------------------------------------------------------------
End Sub
