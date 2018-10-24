Private Sub app(ByVal x As Boolean)

With Application
    If x = False Then
     .StatusBar = "Macro running... Please wait"
    Else
     .StatusBar = Not (x)
    End If
    .ScreenUpdating = x
    .DisplayAlerts = x
End With

End Sub
