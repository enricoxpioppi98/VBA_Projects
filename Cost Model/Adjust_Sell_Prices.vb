Sub AdjustSellPrices()

    Improve_Execution.ScreenUpdating
    
    Set Current_Tab = ActiveSheet

    Sheet_Protection.OFF
    
    'Status Bar -------------------------------------------------------------------------------------------------------------------------------------
        Old_Status_Bar = Application.StatusBar
        Application.StatusBar = True
    '------------------------------------------------------------------------------------------------------------------------------------------------
    
    'Timer ------------------------------------------------------------------------------------------------------------------------------------------
        Dim StartTime As Double
        Dim SecondsElapsed As Double
        
        StartTime = Timer
    '------------------------------------------------------------------------------------------------------------------------------------------------

    'Inputs -----------------------------------------------------------------------------------------------------------------------------------------
        ActiveWorkbook.Worksheets("Financials by Part").Activate

        Profit_Address = "A1"
        Sell_Price_Address = "A1"
        Target_Sell_Price_Address = "A1"
    '------------------------------------------------------------------------------------------------------------------------------------------------

    'Adjust Sell Prices -----------------------------------------------------------------------------------------------------------------------------
        Application.StatusBar = "Adjusting Sell Prices"
        
        For Part_Section_In_FinancialsByPart = 1 To 18

            Profit_Address = "B" & Columns(1).Find(What:="Profit", After:=Cells(Range(Profit_Address).Row, 1), LookAt:=xlWhole).Row
            Sell_Price_Address = "C" & Columns(1).Find(What:="Sell Price ", After:=Cells(Range(Sell_Price_Address).Row, 1), LookAt:=xlWhole).Row
            Target_Sell_Price_Address = "R" & Columns(15).Find(What:="COMMERCIAL ISSUE NOTE ON THIS PART:", After:=Cells(Range(Target_Sell_Price_Address).Row, 15), LookAt:=xlWhole).Row

            Sell_Price = Range(Sell_Price_Address).Value
            Target_Sell_Price = Range(Target_Sell_Price_Address).Value

            If Sell_Price <> 0 And Target_Sell_Price <> 0 Then
                Range(Sell_Price_Address).GoalSeek Goal:=Target_Sell_Price, ChangingCell:=Range(Profit_Address)
            End If
        Next Part_Section_In_FinancialsByPart
    '------------------------------------------------------------------------------------------------------------------------------------------------

    'SG&A -------------------------------------------------------------------------------------------------------------------------------------------
        Application.StatusBar = "Adjusting SGA"
        
        ActiveWorkbook.Worksheets("Cash Flow Forecast").Activate

        Total_SGA_Row = Range("A:E").Find(What:="Total SG&A", After:=Cells(1, 1), LookAt:=xlWhole).Row
        Total_SGA_Address = Split(Cells(1, ActiveSheet.Rows(Total_SGA_Row).Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column).Address, "$")(1) & Total_SGA_Row

        SGA_To_Change_Address = "G" & (Total_SGA_Row - 1)

        Range(Total_SGA_Address).NumberFormat = "0.000%"
        If SGA_To_Change_Address <> "" And SGA_To_Change_Address <> 0 Then
            While Abs(Range(Total_SGA_Address).Value - 0.15) > 0.0001
                SecondsElapsed = Round(Timer - StartTime, 2)
                If SecondsElapsed > 30 Then
                    Improve_Execution.Restore
                    MsgBox "Run time error - Code execution took too long, please check your sell prices, Total SGA %, and SGA - Corporate %"
                    Exit Sub
                End If
                Range(SGA_To_Change_Address).Value = 0
                Range(Total_SGA_Address).GoalSeek Goal:=0.15, ChangingCell:=Range(SGA_To_Change_Address)
            Wend
        End If
    '------------------------------------------------------------------------------------------------------------------------------------------------
    
    Current_Tab.Activate

    Improve_Execution.Restore

End Sub