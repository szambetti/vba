Attribute VB_Name = "Functions"
'restricts to file or folder
Public Enum fileorfolder
    file = 0
    folder = 1
End Enum
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
        .statusbar = False
    End If
    .ScreenUpdating = x
    .DisplayAlerts = x
    DoEvents
End With

End Sub
'checks if a file or a folder exists, if so returns true
Public Function exist(ByVal x As String, ByVal y As fileorfolder) As Boolean

Select Case y
  Case file
    If Len(Dir(x)) > 0 Then
        exist = True
    End If
  Case folder
    If Len(Dir(x, vbDirectory)) > 0 Then
        exist = True
    End If
  Case Else
    Debug.Print "Error on exist function"
End Select

End Function
'+= var
Public Function Inc(ByRef x As Integer) As Integer
   x = x + 1
End Function

