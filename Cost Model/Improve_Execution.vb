Sub ScreenUpdating()

    Application.ScreenUpdating = False

End Sub
Sub ScreenUpdating_And_Calculation()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

End Sub
Sub Restore()

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub
