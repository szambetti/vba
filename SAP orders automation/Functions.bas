Attribute VB_Name = "Functions"
'public enums
Public Enum fileorfolder
    file = 0
    folder = 1
End Enum

Public Enum lyorlastmonth
    ly = 0
    lastmonth = 1
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
''gets ly period correctly
'Public Function getPeriod(ByRef x, Optional lyorlastmonth As String) As String
'
'    Dim tableperiod As ListObject, period_array As Variant, b, j As Integer, ctrl As Worksheet, conv, k As String
'
'    'converts x type into string
'    If conv = TypeName(x) = "Range" Then
'        k = x.Value
'    Else
'        k = x
'    End If
'
'
'    Set ctrl = ThisWorkbook.Worksheets("control panel")
'    Set tableperiod = ctrl.ListObjects("periodtable")
'    period_array = tableperiod.DataBodyRange
'
'    If lyorlastmonth = "ly" Then
'        j = 12
'    ElseIf lyorlastmonth = "lastmonth" Then
'        j = 1
'    End If
'
'    For b = LBound(period_array) To UBound(period_array)
'        If period_array(b, 1) = k Then
'            getPeriod = period_array(b - j, 1)
'            Exit For
'        End If
'    Next
'
'End Function
