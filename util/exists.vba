'checks if a file or a folder exists, if so returns true
Function exist(ByVal x As String, ByVal type as string) As Boolean

Select case type
  Case "file"
    If Len(Dir(x)) > 0 Then
        exist = True
    End If
  Case "folder"
    If Len(Dir(x, vbDirectory)) > 0 Then
        exist = True
    End If
  Case Else
    exists = "error"
End Select

End Function
