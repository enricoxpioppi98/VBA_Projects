Sub Toggle_Structure_Protection()

    If ActiveWorkbook.ProtectStructure = False Then
        ActiveWorkbook.Protect Password:="GCM2016SC"
        MsgBox "Structure protection activated."
    Else
        ActiveWorkbook.Unprotect Password:="GCM2016SC"
        MsgBox "Structure protection unlocked."
    End If

End Sub