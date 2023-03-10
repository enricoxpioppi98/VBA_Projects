Sub Formatting()

    Improve_Execution.ScreenUpdating
    Active_ws = ActiveSheet

    'Unprotect

    ActiveWorkbook.Worksheets("Input").Activate

    Cells.RowHeight = 18
    Range("F:G,I:Q").ColumnWidth = 14
    Range("C:C").ColumnWidth = 15
    
    'Rename items -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        On Error Resume Next

        ActiveSheet.Shapes.Range(Array("Grafik 4")).Name = "Gentherm Logo"
        ActiveSheet.Shapes.Range(Array("Drop Down 46")).Name = "Currency"
        ActiveSheet.Shapes.Range(Array("Drop Down 96")).Name = "New Customer"
        ActiveSheet.Shapes.Range(Array("Drop Down 99")).Name = "New Technology"
        ActiveSheet.Shapes.Range(Array("Drop Down 94")).Name = "License Fee"
        ActiveSheet.Shapes.Range(Array("Drop Down 93")).Name = "Commission"
        ActiveSheet.Shapes.Range(Array("Drop Down 91")).Name = "Plant #18"
        ActiveSheet.Shapes.Range(Array("Drop Down 90")).Name = "Plant #17"
        ActiveSheet.Shapes.Range(Array("Drop Down 89")).Name = "Plant #16"
        ActiveSheet.Shapes.Range(Array("Drop Down 88")).Name = "Plant #15"
        ActiveSheet.Shapes.Range(Array("Drop Down 87")).Name = "Plant #14"
        ActiveSheet.Shapes.Range(Array("Drop Down 86")).Name = "Plant #13"
        ActiveSheet.Shapes.Range(Array("Drop Down 85")).Name = "Plant #12"
        ActiveSheet.Shapes.Range(Array("Drop Down 83")).Name = "Plant #1"
        ActiveSheet.Shapes.Range(Array("Drop Down 81")).Name = "Plant #11"
        ActiveSheet.Shapes.Range(Array("Drop Down 80")).Name = "Plant #10"
        ActiveSheet.Shapes.Range(Array("Drop Down 79")).Name = "Plant #9"
        ActiveSheet.Shapes.Range(Array("Drop Down 78")).Name = "Plant #8"
        ActiveSheet.Shapes.Range(Array("Drop Down 77")).Name = "Plant #7"
        ActiveSheet.Shapes.Range(Array("Drop Down 76")).Name = "Plant #6"
        ActiveSheet.Shapes.Range(Array("Drop Down 75")).Name = "Plant #5"
        ActiveSheet.Shapes.Range(Array("Drop Down 74")).Name = "Plant #4"
        ActiveSheet.Shapes.Range(Array("Drop Down 73")).Name = "Plant #3"
        ActiveSheet.Shapes.Range(Array("Drop Down 72")).Name = "Plant #2"
        ActiveSheet.Shapes.Range(Array("Drop Down 69")).Name = "Segment #18"
        ActiveSheet.Shapes.Range(Array("Drop Down 68")).Name = "Segment #17"
        ActiveSheet.Shapes.Range(Array("Drop Down 67")).Name = "Segment #16"
        ActiveSheet.Shapes.Range(Array("Drop Down 66")).Name = "Segment #15"
        ActiveSheet.Shapes.Range(Array("Drop Down 65")).Name = "Segment #14"
        ActiveSheet.Shapes.Range(Array("Drop Down 64")).Name = "Segment #13"
        ActiveSheet.Shapes.Range(Array("Drop Down 63")).Name = "Segment #12"
        ActiveSheet.Shapes.Range(Array("Drop Down 62")).Name = "Segment #11"
        ActiveSheet.Shapes.Range(Array("Drop Down 61")).Name = "Segment #10"
        ActiveSheet.Shapes.Range(Array("Drop Down 60")).Name = "Segment #9"
        ActiveSheet.Shapes.Range(Array("Drop Down 59")).Name = "Amtz Part #18"
        ActiveSheet.Shapes.Range(Array("Drop Down 58")).Name = "Amtz Part #17"
        ActiveSheet.Shapes.Range(Array("Drop Down 57")).Name = "Amtz Part #16"
        ActiveSheet.Shapes.Range(Array("Drop Down 56")).Name = "Amtz Part #15"
        ActiveSheet.Shapes.Range(Array("Drop Down 55")).Name = "Amtz Part #14"
        ActiveSheet.Shapes.Range(Array("Drop Down 54")).Name = "Amtz Part #13"
        ActiveSheet.Shapes.Range(Array("Drop Down 53")).Name = "Amtz Part #12"
        ActiveSheet.Shapes.Range(Array("Drop Down 52")).Name = "Amtz Part #11"
        ActiveSheet.Shapes.Range(Array("Drop Down 51")).Name = "Amtz Part #10"
        ActiveSheet.Shapes.Range(Array("Drop Down 50")).Name = "Amtz Part #9"
        ActiveSheet.Shapes.Range(Array("Drop Down 49")).Name = "Amtz Year #8"
        ActiveSheet.Shapes.Range(Array("Drop Down 48")).Name = "Amtz Year #7"
        ActiveSheet.Shapes.Range(Array("Check Box 47")).Name = "LTA Non-Material Only"
        ActiveSheet.Shapes.Range(Array("Drop Down 45")).Name = "Terms of Delivery"
        ActiveSheet.Shapes.Range(Array("Drop Down 33")).Name = "Amtz Year #9"
        ActiveSheet.Shapes.Range(Array("Drop Down 32")).Name = "Amtz Year #6"
        ActiveSheet.Shapes.Range(Array("Drop Down 31")).Name = "Quote Type"
        ActiveSheet.Shapes.Range(Array("Drop Down 32")).Name = "Amtz Year #6"
        ActiveSheet.Shapes.Range(Array("Drop Down 29")).Name = "Amtz Part #8"
        ActiveSheet.Shapes.Range(Array("Drop Down 28")).Name = "Amtz Part #7"
        ActiveSheet.Shapes.Range(Array("Drop Down 27")).Name = "Amtz Part #6"
        ActiveSheet.Shapes.Range(Array("Drop Down 26")).Name = "Amtz Part #5"
        ActiveSheet.Shapes.Range(Array("Drop Down 25")).Name = "Amtz Part #4"
        ActiveSheet.Shapes.Range(Array("Drop Down 24")).Name = "Amtz Part #3"
        ActiveSheet.Shapes.Range(Array("Drop Down 23")).Name = "Amtz Part #2"
        ActiveSheet.Shapes.Range(Array("Drop Down 22")).Name = "Amtz Part #1"
        ActiveSheet.Shapes.Range(Array("Drop Down 21")).Name = "Segment #8"
        ActiveSheet.Shapes.Range(Array("Drop Down 20")).Name = "Segment #7"
        ActiveSheet.Shapes.Range(Array("Drop Down 19")).Name = "Segment #6"
        ActiveSheet.Shapes.Range(Array("Drop Down 18")).Name = "Segment #5"
        ActiveSheet.Shapes.Range(Array("Drop Down 17")).Name = "Segment #4"
        ActiveSheet.Shapes.Range(Array("Drop Down 16")).Name = "Segment #3"
        ActiveSheet.Shapes.Range(Array("Drop Down 15")).Name = "Segment #2"
        ActiveSheet.Shapes.Range(Array("Drop Down 7")).Name = "Segment #1"
        ActiveSheet.Shapes.Range(Array("Drop Down 6")).Name = "Customer LTA"
        ActiveSheet.Shapes.Range(Array("Drop Down 5")).Name = "Amtz Year #5"
        ActiveSheet.Shapes.Range(Array("Drop Down 4")).Name = "Amtz Year #4"
        ActiveSheet.Shapes.Range(Array("Drop Down 3")).Name = "Amtz Year #3"
        ActiveSheet.Shapes.Range(Array("Drop Down 2")).Name = "Amtz Year #2"
        ActiveSheet.Shapes.Range(Array("Drop Down 1")).Name = "Amtz Year #1"

        On Error GoTo 0
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Height -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ActiveSheet.Shapes.SelectAll
        Selection.Height = 15.84
        ActiveSheet.Shapes.Range(Array("Gentherm Logo")).Height = 41.04
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Individual Fields --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Dim Individual_Fields(1 To 9) As String
        Individual_Fields(1) = "Currency"
        Individual_Fields(2) = "New Customer"
        Individual_Fields(3) = "New Technology"
        Individual_Fields(4) = "Terms of Delivery"
        Individual_Fields(5) = "Quote Type"
        Individual_Fields(6) = "Customer LTA"
        Individual_Fields(7) = "LTA Non-Material Only"
        Individual_Fields(8) = "Commission"
        Individual_Fields(9) = "License Fee"
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    For Each sh In ActiveSheet.Shapes
        If sh.Name <> "Gentherm Logo" Then
            If InStr(1, "'" & Join(Individual_Fields, "'") & "'", "'" & sh.Name & "'") > 0 Then
                Select Case sh.Name
                Case Is = "Currency"
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Left = Range("C4").Left
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Top = Range("C4").Top
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Width = Application.InchesToPoints(1)
                Case Is = "New Customer"
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Left = Range("C12").Left
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Top = Range("C12").Top
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Width = Application.InchesToPoints(1)
                Case Is = "New Technology"
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Left = Range("C13").Left
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Top = Range("C13").Top
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Width = Application.InchesToPoints(1)
                Case Is = "Terms of Delivery"
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Left = Range("C21").Left
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Top = Range("C21").Top
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Width = Application.InchesToPoints(1)
                Case Is = "Quote Type"
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Width = Application.InchesToPoints(2)
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Left = Range("C30").Left
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Top = Range("C30").Top
                Case Is = "Customer LTA"
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Left = Range("B83").Left
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Top = Range("B83").Top
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Width = Application.InchesToPoints(2)
                Case Is = "LTA Non-Material Only"
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Left = Range("D83").Left
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Top = Range("D83").Top
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Width = Application.InchesToPoints(2)
                Case Is = "Commission"
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Left = Range("B84").Left
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Top = Range("B84").Top
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Width = Application.InchesToPoints(2)
                Case Is = "License Fee"
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Left = Range("B85").Left
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Top = Range("B85").Top
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Width = Application.InchesToPoints(2)
                End Select
            Else
                ActiveSheet.Shapes.Range(Array(sh.Name)).Width = Application.InchesToPoints(1)

                If InStr(sh.Name, "Plant") > 0 Or InStr(sh.Name, "Segment") > 0 Then
                    ActiveSheet.Shapes.Range(Array(sh.Name)).Top = Range("F" & (Split(sh.Name, "#")(1) + 34)).Top
                    If InStr(sh.Name, "Plant") > 0 Then
                        ActiveSheet.Shapes.Range(Array(sh.Name)).Left = Range("F" & (Split(sh.Name, "#")(1) + 34)).Left
                    Else
                        ActiveSheet.Shapes.Range(Array(sh.Name)).Left = Range("G" & (Split(sh.Name, "#")(1) + 34)).Left
                    End If
                End If
                If InStr(sh.Name, "Amtz") > 0 Then
                    If InStr(sh.Name, "Year") > 0 Then
                        ActiveSheet.Shapes.Range(Array(sh.Name)).Top = Range("I56").Top
                        ActiveSheet.Shapes.Range(Array(sh.Name)).Left = Columns(8 + Split(sh.Name, "#")(1)).Left
                    Else
                        ActiveSheet.Shapes.Range(Array(sh.Name)).Left = Range("G60").Left
                        ActiveSheet.Shapes.Range(Array(sh.Name)).Top = Range("G" & (Split(sh.Name, "#")(1) + 59)).Top
                    End If
                End If
            End If
            Number_Of_Shapes = Number_Of_Shapes + 1
        End If
    Next
    
    ActiveSheet.Shapes.Range(Array("Amtz Year #1", _
        "Amtz Year #2", "Amtz Year #3", "Amtz Year #4", "Amtz Year #5", _
        "Segment #1", "Segment #2", "Segment #3", "Segment #4", "Segment #5", _
        "Segment #6", "Segment #7", "Segment #8", "Amtz Part #1", "Amtz Part #2", _
        "Amtz Part #3", "Amtz Part #4", "Amtz Part #5", "Amtz Part #6", "Amtz Part #7" _
        , "Amtz Part #8", "Amtz Year #6", "Amtz Year #9", _
        "Amtz Year #7", "Amtz Year #8", "Amtz Part #9" _
        , "Amtz Part #10", "Amtz Part #11", "Amtz Part #12", "Amtz Part #13", _
        "Amtz Part #14", "Amtz Part #15", "Amtz Part #16", "Amtz Part #17", _
        "Amtz Part #18", "Segment #9", "Segment #10", "Segment #11", "Segment #12", _
        "Segment #13", "Segment #14", "Segment #15", "Segment #16", "Segment #17", _
        "Segment #18", "Plant #2", "Plant #3", "Plant #4", "Plant #5", "Plant #6", _
        "Plant #7", "Plant #8", "Plant #9", "Plant #10", "Plant #11", "Plant #1", _
        "Plant #12", "Plant #13", "Plant #14", "Plant #15", "Plant #16", "Plant #17", _
        "Plant #18")). _
        Select
    Selection.Placement = xlMoveAndSize
    
        ActiveSheet.Shapes.Range(Array("LTA Non-Material Only", "Customer LTA", _
        "Quote Type", _
        "Terms of Delivery", "Currency", _
        "Commission", "License Fee", "New Customer", "New Technology")). _
        Select
    Selection.Placement = xlMove

    'Financials by Part -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ActiveWorkbook.Worksheets("Financials by Part").Activate

        With Cells.SpecialCells(xlCellTypeVisible)
            .EntireRow.AutoFit
        End With

        Target_Sell_Price_Address = "A1"

        For Part_Section_In_FinancialsByPart = 1 To 18

            Target_Sell_Price_Address = "R" & Columns(15).Find(What:="COMMERCIAL ISSUE NOTE ON THIS PART:", After:=Cells(Range(Target_Sell_Price_Address).Row, 15), LookAt:=xlWhole).Row
            Part_Section_Address = "O" & (Range(Target_Sell_Price_Address).Row - 1)

            With Range(Target_Sell_Price_Address)
                .Borders(xlEdgeLeft).LineStyle = xlDouble
                .Borders(xlEdgeTop).LineStyle = xlDouble
                .Borders(xlEdgeBottom).LineStyle = xlDouble
                .Borders(xlEdgeRight).LineStyle = xlDouble
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
                .Interior.Color = 49407
            End With

            With Range(Part_Section_Address)
                .Value = Part_Section_In_FinancialsByPart
                .Font.Bold = True
            End With
        Next i

        'Protect --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            If Was_Protected = 1 Then
                Application.DisplayAlerts = False
                Toggle_Sheet_Protection.Toggle_Sheet_Protection
                Application.DisplayAlerts = True
            End If
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Active_ws.Activate
    Improve_Execution.Restore

End Sub
