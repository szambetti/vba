'checks if a file or a folder exists, if so returns true
Function exist(ByVal x As String) As Boolean

If x Like "*.*" Then
    If Len(Dir(x)) > 0 Then
        exist = True
    End If
Else
    If Len(Dir(x, vbDirectory)) > 0 Then
        exist = True
    End If
End If

End Function
