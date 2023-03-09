Sub Create_Variance_Baseline()

    'Wrong Sheet ------------------------------------------------------------------------------------------------------------------------------------
        Dim Forbidden_Sheets() As String
        Forbidden_Sheets(1) = "Definitions_Decisions"
        foRBidden_Sheets(2) = "Executive Summary"
        Forbidden_Sheets(3) = "Assumptions"
        Forbidden_Sheets(4) = "Business Award Approval"
        Forbidden_Sheets(5) = "Customer Contract Review"
        Forbidden_Sheets(6) = "Contribution Margin"
        Forbidden_Sheets(7) = "Cash Flow Forecast"
        Forbidden_Sheets(8) = "Cost Structure"
        Forbidden_Sheets(9) = "Checklist - Completion of Cost"
        Forbidden_Sheets(10) = "Input"
        Forbidden_Sheets(11) = "Program Summary"
        Forbidden_Sheets(12) = "Financials by Part"
        Forbidden_Sheets(13) = "Prog Costs"
        Forbidden_Sheets(14) = "Freight"
        Forbidden_Sheets(15) = "Capacity"
        Forbidden_Sheets(16) = "Machine Rate"
        Forbidden_Sheets(17) = "Table"
        Forbidden_Sheets(18) = "MX"
        Forbidden_Sheets(19) = "BLW SCRAP%"
        Forbidden_Sheets(20) = "CN"

        'If the name of the current sheet is in the array, then exit the macro
        For Array_i = 1 To Length(Forbidden_Sheets)
            If InStr(ActiveSheet.Name, Forbidden_Sheets(Array_i)) > 0 Then
                MsgBox "This macro cannot be run on this sheet."
                Exit Sub
            End If
        Next Array_i
    '------------------------------------------------------------------------------------------------------------------------------------------------

    'Improve Execution ------------------------------------------------------------------------------------------------------------------------------
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
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
    
    'Restore Execution ------------------------------------------------------------------------------------------------------------------------------
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    '------------------------------------------------------------------------------------------------------------------------------------------------
    
End Sub
