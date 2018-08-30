VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'As soon as the workbook is opened, it checks for the Analysis COMAddIn.
'If it finds Analysis and thinks it isn’t enabled, it enables it.
'If it finds Analysis and thinks its already enabled, it disables it first then re-enables it to make doubly sure its going to function ok.'

Private Sub Workbook_Open()
Application.statusbar = "Enabling analysis... please wait"

ThisWorkbook.Worksheets("Daily_Tables").Activate
range("A1").Select
range("G4") = "Enabling analysis..."
range("D5") = 0
DoEvents
Application.ScreenUpdating = False

Dim lResult As Long

    Dim addin As COMAddIn

    For Each addin In Application.COMAddIns

        If addin.progID = "SapExcelAddIn" Then

            If addin.Connect = False Then
                On Error GoTo continue
                addin.Connect = True
                On Error GoTo -1
                On Error GoTo 0
            ElseIf addin.Connect = True Then
                addin.Connect = False
                addin.Connect = True
            End If

        End If

    Next
Application.ScreenUpdating = True

range("G4") = "Idle"
range("d5") = 0.05
Application.statusbar = False

iinput = MsgBox("Would you like to start the generation procedure?", vbYesNo)
If iinput = vbYes Then
    Call main
    Exit Sub
Else
End If
Exit Sub

continue:

range("G4") = "Idle"
range("d5") = 0.05
Application.statusbar = False
MsgBox ("Error, analysis could not be enabled automatically... please do so manually")

End Sub
