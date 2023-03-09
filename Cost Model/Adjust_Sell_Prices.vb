Sub AdjustSellPrices()

    'Improve Executiion -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Application.ScreenUpdating = False
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'WorkBooks ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Set Current_Tab = ActiveSheet
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'Inputs -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ActiveWorkbook.Worksheets("Financials by Part").Activate

        Current_SellPrice_Address = "C" & ActiveSheet.Columns(1).Find("Sell Price").Row
        Target_SellPrice_Address = "R40"
        Profit_Address = "B" & ActiveSheet.Columns(1).Find("Profit").Row
        FinancialsByPart_SectionHeight = ActiveSheet.Range("C30:M90").Find("INTERNAL").Row - ActiveSheet.Range("C1:M30").Find("INTERNAL").Row
        ActiveWorkbook.Worksheets("Cash Flow Forecast").Activate
        SGA_Address_Row = ActiveSheet.Range("A:E").Find("Total SG&A").Row
        SGA_Address = Split(Cells(1, ActiveSheet.Rows(SGA_Address_Row).Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column).Address, "$")(1) & SGA_Address_Row
        SGA_ToChange_Address = "G" & (SGA_Address_Row - 1)
        ActiveWorkbook.Worksheets("Financials by Part").Activate

        Current_sellPrice_Row = Range(Current_SellPrice_Address).Row
        Current_sellPrice_Column = Range(Current_SellPrice_Address).Column

        Target_SellPrice_Row = Range(Target_SellPrice_Address).Row
        Target_SellPrice_Column = Range(Target_SellPrice_Address).Column

        Profit_Row = Range(Profit_Address).Row
        Profit_Column = Range(Profit_Address).Column
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'Sell Prices --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        For xPart = 1 To 18
            Current_SellPrice = Worksheets("Financials by Part").Cells(Current_sellPrice_Row + (FinancialsByPart_SectionHeight * (xPart - 1)), Current_sellPrice_Column).Value
            Target_SellPrice = Worksheets("Financials by Part").Cells(Target_SellPrice_Row + (FinancialsByPart_SectionHeight * (xPart - 1)), Target_SellPrice_Column).Value
            Current_SellPrice_Address_NEW = Split(Cells(1, Current_sellPrice_Column).Address, "$")(1) & (Current_sellPrice_Row + (FinancialsByPart_SectionHeight * (xPart - 1)))
            Profit_Address_NEW = Split(Cells(1, Profit_Column).Address, "$")(1) & (Profit_Row + (FinancialsByPart_SectionHeight * (xPart - 1)))
            If Current_SellPrice <> 0 And Current_SellPrice <> "-" And Target_SellPrice <> "" And Target_SellPrice <> 0 Then
                Range(Current_SellPrice_Address_NEW).GoalSeek Goal:=Target_SellPrice, ChangingCell:=Range(Profit_Address_NEW)
            End If
        Next xPart
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'SG&A ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        xRepeat:
        If SGA_ToChange_Address <> "" And SGA_ToChange_Address <> 0 Then
            Worksheets("Cash Flow Forecast").Activate
            Range(SGA_Address).NumberFormat = "0.000%"
            Range(SGA_Address).GoalSeek Goal:=0.15, ChangingCell:=Range(SGA_ToChange_Address)
        End If
        If Abs(Range(SGA_Address).Value - 0.15) > 0.0001 Then
            Range(SGA_ToChange_Address).Value = 0
            GoTo xRepeat
        End If
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Current_Tab.Activate

    'Restore Execution --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Application.ScreenUpdating = True
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub