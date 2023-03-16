Sub Toggle()

    Improve_Execution.ScreenUpdating_And_Calculation
    Set Active_ws = ActiveSheet

    'Tabs With Long Password ------------------------------------------------------------------------------------------------------------------------
        Dim Parts_With_Long_Password(1 To 20) As String

        Parts_With_Long_Password(1) = "Executive Summary-ROI"
        Parts_With_Long_Password(2) = "Assumptions"
        Parts_With_Long_Password(3) = "Business Award Approval - DOA"
        Parts_With_Long_Password(4) = "Customer Contract Review"
        Parts_With_Long_Password(5) = "Contribution Margin"
        Parts_With_Long_Password(6) = "Cash Flow Forecast"
        Parts_With_Long_Password(7) = "Cost Structure"
        Parts_With_Long_Password(8) = "Table"
    '------------------------------------------------------------------------------------------------------------------------------------------------

    'Tabs With Short Password -----------------------------------------------------------------------------------------------------------------------
        Dim Parts_With_Short_Password(1 To 20) As String

        Parts_With_Short_Password(1) = "Definitions_Decisions"
        Parts_With_Short_Password(2) = "Input"
        Parts_With_Short_Password(3) = "Program Summary"
        Parts_With_Short_Password(4) = "Financials By Part"
        Parts_With_Short_Password(5) = "Prog Costs"
        Parts_With_Short_Password(6) = "Freight"
        Parts_With_Short_Password(7) = "Capacity"
        Parts_With_Short_Password(8) = "Machine Rate"
        Parts_With_Short_Password(9) = "CN"
        Parts_With_Short_Password(10) = "MX"
        Parts_With_Short_Password(11) = "NMK"
        Parts_With_Short_Password(12) = "UA"
        Parts_With_Short_Password(13) = "VN"
    '------------------------------------------------------------------------------------------------------------------------------------------------

    'Set Toggle_Protect_To Flag ----------------------------------------------------------------------------------------------------------------------------
        Dim Toggle_Protect_To As Boolean

        If Active_ws.ProtectContents = False Then
            Toggle_Protect_To = True
        Else
            Toggle_Protect_To = False
        End If
    '------------------------------------------------------------------------------------------------------------------------------------------------

    'Toggle Protection ------------------------------------------------------------------------------------------------------------------------------
        If Toggle_Protect_To = True Then
            For Each ws In ActiveWorkbook.Worksheets
                If ws.ProtectContents = False Then
                    'Check Password Length ----------------------------------------------------------------------------------------------------------
                        If Custom_Function.IsInArray_1D(ws.Name, Parts_With_Long_Password) > 0 Then
                            Tab_Password_Length = "Long"
                        ElseIf Custom_Function.IsInArray_1D(ws.Name, Parts_With_Short_Password) > 0 Then
                            Tab_Password_Length = "Short"
                        Else
                            Tab_Password_Length = "Short"
                        End If
                    '----------------------------------------------------------------------------------------------------------------------------------
                    If Tab_Password_Length = "Long" Then
                        ws.Protect "GCM2016EconCalc"
                    Else
                        ws.Protect "GCM2016SC"
                    End If
                End If
            Next ws
        Else
            For Each ws In ActiveWorkbook.Worksheets
                If ws.ProtectContents = True Then
                    'Check Password Length ----------------------------------------------------------------------------------------------------------
                        If Custom_Function.IsInArray_1D(ws.Name, Parts_With_Long_Password) > 0 Then
                            Tab_Password_Length = "Long"
                        ElseIf Custom_Function.IsInArray_1D(ws.Name, Parts_With_Short_Password) > 0 Then
                            Tab_Password_Length = "Short"
                        Else
                            Tab_Password_Length = "Short"
                        End If
                    '----------------------------------------------------------------------------------------------------------------------------------
                    If Tab_Password_Length = "Long" Then
                        ws.Unprotect "GCM2016EconCalc"
                    Else
                        ws.Unprotect "GCM2016SC"
                    End If
                End If
            Next ws
        End If
    '------------------------------------------------------------------------------------------------------------------------------------------------

    Active_ws.Activate
    Improve_Execution.Restore
    
    If Active_ws.ProtectContents = True Then
        MsgBox "Sheet Protection enabled."
    Else
        MsgBox "Sheet Protection unlocked."
    End If

End Sub

Sub Enable()

    Set Active_ws = ActiveSheet

    'Tabs With Long Password ------------------------------------------------------------------------------------------------------------------------
        Dim Parts_With_Long_Password(1 To 20) As String

        Parts_With_Long_Password(1) = "Executive Summary-ROI"
        Parts_With_Long_Password(2) = "Assumptions"
        Parts_With_Long_Password(3) = "Business Award Approval - DOA"
        Parts_With_Long_Password(4) = "Customer Contract Review"
        Parts_With_Long_Password(5) = "Contribution Margin"
        Parts_With_Long_Password(6) = "Cash Flow Forecast"
        Parts_With_Long_Password(7) = "Cost Structure"
        Parts_With_Long_Password(8) = "Table"
    '------------------------------------------------------------------------------------------------------------------------------------------------

    'Tabs With Short Password -----------------------------------------------------------------------------------------------------------------------
        Dim Parts_With_Short_Password(1 To 20) As String

        Parts_With_Short_Password(1) = "Definitions_Decisions"
        Parts_With_Short_Password(2) = "Input"
        Parts_With_Short_Password(3) = "Program Summary"
        Parts_With_Short_Password(4) = "Financials By Part"
        Parts_With_Short_Password(5) = "Prog Costs"
        Parts_With_Short_Password(6) = "Freight"
        Parts_With_Short_Password(7) = "Capacity"
        Parts_With_Short_Password(8) = "Machine Rate"
        Parts_With_Short_Password(9) = "CN"
        Parts_With_Short_Password(10) = "MX"
        Parts_With_Short_Password(11) = "NMK"
        Parts_With_Short_Password(12) = "UA"
        Parts_With_Short_Password(13) = "VN"
    '------------------------------------------------------------------------------------------------------------------------------------------------

    'Enable Protection ------------------------------------------------------------------------------------------------------------------------------
        For Each ws In ActiveWorkbook.Worksheets
            If ws.ProtectContents = False Then
                'Check Password Length ----------------------------------------------------------------------------------------------------------
                    If Custom_Function.IsInArray_1D(ws.Name, Parts_With_Long_Password) > 0 Then
                        Tab_Password_Length = "Long"
                    ElseIf Custom_Function.IsInArray_1D(ws.Name, Parts_With_Short_Password) > 0 Then
                        Tab_Password_Length = "Short"
                    Else
                        Tab_Password_Length = "Short"
                    End If
                '----------------------------------------------------------------------------------------------------------------------------------
                If Tab_Password_Length = "Long" Then
                    ws.Protect "GCM2016EconCalc"
                Else
                    ws.Protect "GCM2016SC"
                End If
            End If
        Next ws
    '------------------------------------------------------------------------------------------------------------------------------------------------

    Active_ws.Activate

End Sub
Sub OFF()

    Set Active_ws = ActiveSheet

    On Error Resume Next
    For Each ws In ActiveWorkbook.Worksheets
        If ws.ProtectContents = True Then
            ws.Unprotect "GCM2016EconCalc"
            ws.Unprotect "GCM2016SC"
        End If
    Next ws
    On Error GoTo 0

    Active_ws.Activate

End Sub
Sub Toggle_Structure()

    Improve_Execution.ScreenUpdating_And_Calculation

    If ActiveWorkbook.ProtectStructure = False Then
        ActiveWorkbook.Protect Password:="GCM2016SC"
        MsgBox "Structure protection activated."
    Else
        ActiveWorkbook.Unprotect Password:="GCM2016SC"
        MsgBox "Structure protection unlocked."
    End If
    
    Improve_Execution.Restore

End Sub
Private Sub Enable_Structure_Before_Close(Cancel As Boolean)

    If ThisWorkbook.ProtectStructure = False Then
        ThisWorkbook.Protect Password:="GCM2016SC"
    End If
    
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Tab.Color = 2251504 Then
            pw = "GCM2016EconCalc"
        Else
            pw = "GCM2016SC"
        End If
        If ws.ProtectContents = False Then
            ws.Protect pw
        End If
    Next ws

End Sub

