Public Function Is_A_Part_Tab(ws As Worksheet, wb As Workbook, Return_To_Current_Tab As Boolean) As Boolean

    'This macro returns true if the passed worksheet is a tab for a Part in the Cost Model, false otherwise
    
    Dim Part_Tabs_Names(1 To 18) As String
    
    Set Active_ws = ActiveSheet

    'Pull List of Part Tabs From Financials By Part -------------------------------------------------------------------------------------------------
        wb.Worksheets("Financials By Part").Activate
        Material_Cost_Address = "A1"

        For Part_Section_In_FinancialsByPart = 1 To 18
            
            Material_Cost_Address = "C" & ActiveSheet.Columns(1).Find(What:="Material Cost", After:=Cells(Range(Material_Cost_Address).Row, 1), LookAt:=xlWhole).Row

            Formula_To_Split = Range(Material_Cost_Address).Formula
            
            Part_Tabs_Names(Part_Section_In_FinancialsByPart) = Split(Formula_To_Split, "!")(0)
            Part_Tabs_Names(Part_Section_In_FinancialsByPart) = Split(Part_Tabs_Names(Part_Section_In_FinancialsByPart), ">0,")(1)
            
            If Left(Part_Tabs_Names(Part_Section_In_FinancialsByPart), 1) = "'" Then
                Part_Tabs_Names(Part_Section_In_FinancialsByPart) = Right(Part_Tabs_Names(Part_Section_In_FinancialsByPart), Len(Part_Tabs_Names(Part_Section_In_FinancialsByPart)) - 1)
            End If
            If Right(Part_Tabs_Names(Part_Section_In_FinancialsByPart), 1) = "'" Then
                Part_Tabs_Names(Part_Section_In_FinancialsByPart) = Left(Part_Tabs_Names(Part_Section_In_FinancialsByPart), Len(Part_Tabs_Names(Part_Section_In_FinancialsByPart)) - 1)
            End If

        Next Part_Section_In_FinancialsByPart
    '------------------------------------------------------------------------------------------------------------------------------------------------

    'Compare ws.Name to Part_Tabs_Names
        For Part_Section_In_FinancialsByPart = 1 To 18
            If ws.Name = Part_Tabs_Names(Part_Section_In_FinancialsByPart) Then
                Is_A_Part_Tab = True
                If Return_To_Current_Tab = True Then
                    Active_ws.Activate
                End If
                Exit Function
            End If
        Next Part_Section_In_FinancialsByPart
    '------------------------------------------------------------------------------------------------------------------------------------------------
        Is_A_Part_Tab = False
        If Return_To_Current_Tab = True Then
            Active_ws.Activate
        End If
        
End Function

Public Function IsInArray_1D(String_To_Find As String, arr As Variant) As Integer

    'This function returns the index of the string in the array if String_To_Find is found, 0 otherwise

    IsInArray_1D = InStr(Join(arr, ""), String_To_Find)

End Function

