Sub New_Quote()

    Improve_Execution.ScreenUpdating
    Set Active_ws = ActiveSheet
    Sheet_Protection.OFF

    'Input Tab --------------------------------------------------------------------------------------------------------------------------------------
        Worksheets("Input").Activate

        Range("C5:E5").Copy
        Range("C6").PasteSpecial xlPasteAllExceptBorders

        'Change the last letter of the string in C5 to the next letter in the alphabet. If it is Z, add an A to the end of the string.
        If Right(Range("C5"), 1) = "Z" Then
            Range("C5") = Range("C5") & "A"
        Else
            Range("C5") = Left(Range("C5"), Len(Range("C5")) - 1) & Chr(Asc(Right(Range("C5"), 1)) + 1)
        End If

        'In cell E5, add the current date as a string
        Range("E5") = "'" & Format(Date, "dd mmm yyyy")

        'From Row 35 To 52,
        'if a cell has a value of "Part #" and the number of the row - 34
        'hide the row
        For i = 35 To 52
            If Range(Cells(i, 1), Cells(i, 1)).Value = "Part #" & i - 34 Then
                Rows(i).Hidden = True
            End If
        Next i

        'From Row 60 To 77,
        'if a cell has a value of "Part #" and the number of the row - 59
        'hide the row
        For i = 60 To 77
            If Range(Cells(i, 1), Cells(i, 1)).Value = "Part #" & i - 59 Then
                Rows(i).Hidden = True
            End If
        Next i

        'Hide Rows from 89 to 98
        Rows("89:98").Hidden = True

    '------------------------------------------------------------------------------------------------------------------------------------------------

    'Program Summary Tab ----------------------------------------------------------------------------------------------------------------------------
        Worksheets("Program Summary").Activate
        Improve_Execution.ScreenUpdating

        'Starting from column G and repeating every 4 columns until column AM,
        'copy values from rows from 16 to 33 and paste them in the previous column
        For i = 7 To 39 Step 4
            Range(Cells(16, i), Cells(33, i)).Copy
            Range(Cells(16, i - 1), Cells(33, i - 1)).PasteSpecial xlPasteValues
        Next i

        New_Note
        Improve_Execution.ScreenUpdating
        
        'In column A, from row 16 to 33,
        'if a cell has a value of "Part #" and the number of the row - 15
        'hide the row
        For i = 16 To 33
            If Range(Cells(i, 1), Cells(i, 1)).Value = "Part #" & i - 15 Then
                Rows(i).Hidden = True
            End If
        Next i

        'In column A, from row 41 to 59,
        'if a cell has a value of "Part #" and the number of the row - 40
        'hide the row
        For i = 41 To 59
            If Range(Cells(i, 1), Cells(i, 1)).Value = "Part #" & i - 40 Then
                Rows(i).Hidden = True
            End If
        Next i

    '------------------------------------------------------------------------------------------------------------------------------------------------

    'Financials by Part -----------------------------------------------------------------------------------------------------------------------------
        Worksheets("Financials by Part").Activate
        Improve_Execution.ScreenUpdating

        Headers_Row = 1
        Sell_Price_Label_Address = "A1"
        Gross_Margin_Label_Address = "A1"
        Sell_Price_Address = "A1"
        First_Part_Flag = 1

        Columns("P:Z").ColumnWidth = 15

        For Part_Section_In_FinancialsByPart = 1 To 18

            Headers_Row = Columns(15).Find(What:="COMMERCIAL ISSUE NOTE ON THIS PART:", After:=Cells(Range(Sell_Price_Label_Address).Row, 15), LookAt:=xlWhole).Row + 9
            Sell_Price_Label_Address = "O" & Columns(15).Find(What:="COMMERCIAL ISSUE NOTE ON THIS PART:", After:=Cells(Range(Sell_Price_Label_Address).Row, 15), LookAt:=xlWhole).Row + 10
            Sell_Price_Address = "C" & Columns(1).Find(What:="Sell Price ", After:=Cells(Range(Sell_Price_Address).Row, 1), LookAt:=xlWhole).Row
            Gross_Margin_Address = "M" & Range(Sell_Price_Address).Row + 3
            Gross_Margin_Label_Address = "O" & Columns(15).Find(What:="COMMERCIAL ISSUE NOTE ON THIS PART:", After:=Cells(Range(Gross_Margin_Label_Address).Row, 15), LookAt:=xlWhole).Row + 11

            Range(Sell_Price_Label_Address) = "Sell Price"
            Range(Sell_Price_Label_Address).Font.Bold = True
            Range(Sell_Price_Label_Address).HorizontalAlignment = xlRight
            Range("P" & Range(Sell_Price_Label_Address).Row & ":Z" & Range(Sell_Price_Label_Address).Row).Font.Bold = False
            Range("P" & Range(Sell_Price_Label_Address).Row & ":Z" & Range(Sell_Price_Label_Address).Row).NumberFormat = "0.0000"

            Range(Gross_Margin_Label_Address) = "Gross Margin"
            Range(Gross_Margin_Label_Address).Font.Bold = True
            Range(Gross_Margin_Label_Address).HorizontalAlignment = xlRight
            Range("P" & Range(Gross_Margin_Label_Address).Row & ":Z" & Range(Gross_Margin_Label_Address).Row).Font.Bold = False
            Range("P" & Range(Gross_Margin_Label_Address).Row & ":Z" & Range(Gross_Margin_Label_Address).Row).NumberFormat = "0.00%"

            Range("P" & Headers_Row & ":Z" & Headers_Row).Font.Bold = True
            Range("P" & Headers_Row & ":Z" & Headers_Row).HorizontalAlignment = xlCenter

            Range("P" & Headers_Row & ":Z" & Headers_Row + 2).Copy
            Range("P" & Headers_Row).PasteSpecial xlPasteValues

            If First_Part_Flag = 1 Then
                First_Available_Column = Range("O" & Range(Sell_Price_Label_Address).Row & ":Z" & Range(Gross_Margin_Label_Address).Row).Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Column + 1
                First_Part_Flag = 0
            End If
            
            'In row Headers_Row in First_Available_Column, enter the digits after "-" in the string in cell C5 in input tab
            Cells(Headers_Row, First_Available_Column).Value = Split(Worksheets("Input").Range("C5").Value, "-")(1)
            'In the same row as Sell_Price_Label_Address in column First_Available_Column,
            Cells(Range(Sell_Price_Label_Address).Row, First_Available_Column).Formula = "=" & Split(Cells(1, Range(Sell_Price_Address).Column).Address, "$")(1) & Range(Sell_Price_Address).Row
            Cells(Range(Gross_Margin_Label_Address).Row, First_Available_Column).Formula = "=" & Split(Cells(1, Range(Gross_Margin_Address).Column).Address, "$")(1) & Range(Gross_Margin_Address).Row

        Next Part_Section_In_FinancialsByPart

        Improve_Execution.ScreenUpdating
    '------------------------------------------------------------------------------------------------------------------------------------------------

    'Parts ------------------------------------------------------------------------------------------------------------------------------------------
        Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
            If Custom_Function.Is_A_Part_Tab(ws, ActiveWorkbook, True) = True Then
                ws.Activate
                Create_Variance_Baseline
                Improve_Execution.ScreenUpdating
            End If
        Next ws
    '------------------------------------------------------------------------------------------------------------------------------------------------

    'Executive Summary ------------------------------------------------------------------------------------------------------------------------------
        Worksheets("Executive Summary-ROI").Activate

        'Copy the values and format in range L48:L58 and paste them in range R48:R58
        Range("L48:L58").Copy
        Range("R48:R58").PasteSpecial xlPasteValues
        Range("R48:R58").PasteSpecial xlPasteFormats
        
        'In cell R47, add "A" bolded and center aligned
        Range("R47") = "A"
        Range("R47").Font.Bold = True
        Range("R47").HorizontalAlignment = xlCenter

        'In cell S47, add "d" bolded and center aligned, column width = 15
        Range("S47") = "d"
        Range("S47").Font.Bold = True
        Range("S47").HorizontalAlignment = xlCenter
        Columns("S:S").ColumnWidth = 15

        'In range S48:S58 calculate the difference between the cell to the left with the cell 7 columns to theleft
        Range("S48:S58").Formula = "=IFERROR(RC[-7]-RC[-1],0)"

        'In range S48:S58 add light gray solid borders on all sides and light gray dotted lines in the center

        Range("S48:S58").Borders(xlEdgeLeft).LineStyle = xlContinuous
        Range("S48:S58").Borders(xlEdgeLeft).Weight = xlThin
        Range("S48:S58").Borders(xlEdgeLeft).ColorIndex = 15

        Range("S48:S58").Borders(xlEdgeTop).LineStyle = xlContinuous
        Range("S48:S58").Borders(xlEdgeTop).Weight = xlThin
        Range("S48:S58").Borders(xlEdgeTop).ColorIndex = 15

        Range("S48:S58").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("S48:S58").Borders(xlEdgeBottom).Weight = xlThin
        Range("S48:S58").Borders(xlEdgeBottom).ColorIndex = 15

        Range("S48:S58").Borders(xlEdgeRight).LineStyle = xlContinuous
        Range("S48:S58").Borders(xlEdgeRight).Weight = xlThin
        Range("S48:S58").Borders(xlEdgeRight).ColorIndex = 15

        Range("S48:S58").Borders(xlInsideHorizontal).LineStyle = xlDot
        Range("S48:S58").Borders(xlInsideVertical).Weight = xlThin
        Range("S48:S58").Borders(xlInsideVertical).ColorIndex = 15

        'In range S48:S58, set intenral color to light yellow
        Range("S48:S58").Interior.ColorIndex = 36

        'Format cell R48 as Custom "[Color50]_(#,##0_)"?";[Red]_(#,##0_)"?";_("-"??_);_(@_)"
        Range("S48:S50").NumberFormat = " [Color50]_(#,##0_)" & ChrW(&H25B2) & ";[Red]_(#,##0_)" & ChrW(&H25BC) & ";_("" - ""??_);_(@_)"
        Range("S51").NumberFormat = "[Red]_(#,##0_)" & ChrW(&H25B2) & ";[Color50]_(#,##0_)" & ChrW(&H25BC) & ";_("" - ""??_);_(@_)"
        Range("S52").NumberFormat = "[Red]_(0.00%_)" & ChrW(&H25B2) & ";[Color50]_(0.00%_)" & ChrW(&H25BC) & ";_("" - ""??_);_(@_)"
        Range("S53:S58").NumberFormat = "[Color50]_(0.00%_)" & ChrW(&H25B2) & ";[Red]_(0.00%_)" & ChrW(&H25BC) & ";_("" - ""??_);_(@_)"


    '------------------------------------------------------------------------------------------------------------------------------------------------
    
    Active_ws.Activate
    Improve_Execution.Restore

End Sub