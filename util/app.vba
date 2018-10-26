'start/stop application macro, allows custom statusbar as optional
Public Sub freeze(ByVal x As Boolean, Optional ByVal StatusbarStr As String)

'inverts as if freeze is true x must be actually false
Let x = Not (x)

With Application
    If x = False Then
        If StatusbarStr = "" Then
            .statusbar = "Macro running... Please wait"
        Else
            .statusbar = StatusbarStr
        End If
    Else
        .statusbar = Not (x)
    End If
    .ScreenUpdating = x
    .DisplayAlerts = x
    DoEvents
End With

End Sub
