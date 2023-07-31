Sub Create_Variance_Baseline()

    Improve_Execution.ScreenUpdating

    'Wrong Sheet ------------------------------------------------------------------------------------------------------------------------------------
        If Custom_Function.Is_A_Part_Tab(ActiveSheet, ActiveWorkbook, True) = False Then
            Improve_Execution.Restore
            MsgBox "This macro cannot be run on this sheet."
            Exit Sub
        End If
    '------------------------------------------------------------------------------------------------------------------------------------------------

    Range("AK17:AL266, AP17:AQ266, AK275:AL284, AP275:AQ284, AK295:AL364, AP295:AQ364").ClearContents
    
    Improve_Execution.ScreenUpdating_And_Calculation

    'For each cell in the range, if the cell is not blank and not zero, copy the value to the cell two columns to the left.
    For Each cell In Range("AM17:AN266, AR17:AS266, AM275:AN284, AR275:AS284, AM295:AN364, AR295:AS364")
        If cell <> "" And cell <> 0 Then
            cell.Offset(0, -2) = cell
        End If
    Next cell
    
    Improve_Execution.Restore
    
End Sub