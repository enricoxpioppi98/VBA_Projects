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

    'BOM --------------------------------------------------------------------------------------------------------------------------------------------
        Range("AM17:AN266").Copy
        Range("AK17").PasteSpecial Paste:=xlPasteValues

        Range("AR17:AS266").Copy
        Range("AP17").PasteSpecial Paste:=xlPasteValues
    '------------------------------------------------------------------------------------------------------------------------------------------------
    'Packaging --------------------------------------------------------------------------------------------------------------------------------------
        Range("AM275:AN284").Copy
        Range("AK275").PasteSpecial Paste:=xlPasteValues

        Range("AR275:AS284").Copy
        Range("AP275").PasteSpecial Paste:=xlPasteValues
    '------------------------------------------------------------------------------------------------------------------------------------------------
    'Routings ---------------------------------------------------------------------------------------------------------------------------------------
        Range("AK295").PasteSpecial Paste:=xlPasteValues

        Range("AR295:AS364").Copy
        Range("AP295").PasteSpecial Paste:=xlPasteValues
    '------------------------------------------------------------------------------------------------------------------------------------------------
    
    Improve_Execution.Restore
    
End Sub