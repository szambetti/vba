'restricts intellisense to file or folder
Public Enum fileorfolder
    file = 0
    folder = 1
End Enum
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
