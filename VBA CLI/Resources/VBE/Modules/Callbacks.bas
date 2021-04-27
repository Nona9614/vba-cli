Public Sub onLoad(ribbon As IRibbonUI)
    MsgBox "Welcome to VBA!", vbOKOnly + vbInformation, "Debug Mode"
End Sub

Public Sub onDebug(control As IRibbonControl)
    MsgBox "Debugging", vbOKOnly + vbInformation, "Debug Mode"
End Sub