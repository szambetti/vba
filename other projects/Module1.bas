Attribute VB_Name = "Module1"

Sub SplitIntoFiles()

Dim ws, ws1 As Worksheet, yymm, WbStr, wbpath As String
Const xl_ext = ".xlsx"

Let yymm = Format(Date, "yy") & Format(Date, "mm")

Call app(False)

With ThisWorkbook
    .Save
    Let wbpath = .path & "\" & yymm
    If exist(wbpath) = False Then
        MkDir (wbpath)
    End If
    For Each ws In .Worksheets
        If Not (ws.Name = "control panel" Or ws.Name = "Template") Then
            Let WbStr = wbpath & "\" & yymm & "_BluePrint Controlling_" & ws.Name & xl_ext
            ThisWorkbook.Sheets.Copy
            ActiveWorkbook.SaveAs WbStr
                For Each ws1 In ActiveWorkbook.Worksheets
                    If Not ws1.Name = ws.Name Then
                        If ws1.Name = "control panel" Then
                            Worksheets(ws1.Name).Visible = False
                        Else
                        Sheets(ws1.Name).Delete
                        End If
                    End If
                Next ws1
            ActiveWorkbook.Close True
        End If
    Next ws
End With

Call app(True)

MsgBox "Master successfully split into the regions in the " & yymm & " subfolder."

End Sub
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

Private Sub app(ByVal x As Boolean)

With Application
    If x = False Then
     .StatusBar = "Macro running... Please wait"
    Else
     .StatusBar = x
    End If
    .ScreenUpdating = x
    .DisplayAlerts = x
End With

End Sub
