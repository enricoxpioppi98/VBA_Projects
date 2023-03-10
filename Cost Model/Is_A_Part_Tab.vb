Public Function Is_A_Part_Tab(ws As Worksheet, wb As Workbook) As Boolean 

    'This macro returns true if the worksheet is a part tab, false otherwise

    Dim Part_Tabs_Names(1 To 18) As String

    'Pull List of Part Tabs From Financials By Part -------------------------------------------------------------------------------------------------
        wb.Worksheets("Financials By Part").Activate
        Cells(1, 1).Select

        FinancialsByPart_SectionHeight = ActiveSheet.Range("C30:M90").Find("INTERNAL").Row - ActiveSheet.Range("C1:M30").Find("INTERNAL").Row '47

        For Part_Section_In_FinancialsByPart = 1 To 18
            
            ActiveSheet.Columns(1).Find(What:="Material Cost", After:=ActiveCell, LookAt:=xlWhole).Select
            Material_Cost_Address = "C" & Selection.Row

            Formula_To_Split = Range(Material_Cost_Address).Formula

            Part_Tabs_Names(Part_Section_In_FinancialsByPart) = Split(Formula_To_Split, "'")(1)

        Next Part_Section_In_FinancialsByPart
    '------------------------------------------------------------------------------------------------------------------------------------------------


End Function

'42028041A FSCH