'changes quickly alerts, status and screenupdating
Function app(x As Boolean) As Application

With Application
    .ScreenUpdating = x
    .DisplayAlerts = x
    If x = true Then
      .status = "Macro running... Please wait"
    Else
     .status = x
    End If
End With

End Function
