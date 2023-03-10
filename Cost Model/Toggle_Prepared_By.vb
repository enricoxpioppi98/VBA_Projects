Sub Toggle_Prepared_By()
    
    Improve_Execution.ScreenUpdating_And_Calculation
    
    Set Active_Cell = ActiveCell
    
    'Correct Or Wrong Sheet -------------------------------------------------------------------------------------------------------------------------
        If Custom_Function.Is_A_Part_Tab(ActiveSheet, ActiveWorkbook, True) = True Then
            Set Active_ws = ActiveWorkbook.ActiveSheet
            'Check if G12 is empty or contains an "x" and set toogle_ON flag ------------------------------------------------------------------------
                If Active_ws.Range("G12").Value = "" Or Active_ws.Range("G12").Value = "x" Then
                    Toggle_ON = True
                Else
                    Toggle_ON = False
                End If
            '----------------------------------------------------------------------------------------------------------------------------------------
        Else
            Active_Cell.Select
            Improve_Execution.Restore
            MsgBox "This macro cannot be run on this sheet."
            Exit Sub
        End If
    '------------------------------------------------------------------------------------------------------------------------------------------------

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If Custom_Function.Is_A_Part_Tab(ws, ActiveWorkbook, False) = True Then
            If Toggle_ON = True Then
                ws.Range("G12, J271, G291, G369").Value = "Enrico Pioppi"
                ws.Range("G13, J272, G292, G370").Value = Format(Date, "dd-mmm-yyyy")
            Else
                ws.Range("G12, J271, G291, G369, G13, J272, G292, G370").Value = "x"
            End If
        End If
    Next ws
    
    Active_ws.Activate
    Improve_Execution.Restore

End Sub