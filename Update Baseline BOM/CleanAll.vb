Sub CleanAll()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    xCurrentRow = ActiveCell.Row
    xCurrentColumn = ActiveCell.Column
    
    Range("A3:F143361").ClearContents
    
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
    
    Cells(xCurrentRow, xCurrentColumn).Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub
