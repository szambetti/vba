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
ThisWorkbook.Worksheets("Daily_Tables").Activate
[tables_A1].Select

'if template is stored in its template folder then being macro run
'else exit
If ThisWorkbook.Path Like "*Template*" Then

    Application.statusbar = "Enabling analysis... please wait"
    [state_rng] = "Enabling analysis..."
    [progressbar_rng] = 0
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
[state_rng] = "Idle"
[progressbar_rng] = 0.05
Application.statusbar = False
    
    Call start_form(False)
    
    If uinput = 0 Then
        Call start_form(True)
        Call main
    ElseIf uinput = 1 Then
        Call start_form(True)
    Else
        Call start_form(True)
        Call month_settings_display
        DoEvents
        Call main
        Exit Sub
    End If
Exit Sub
continue:
    [state_rng] = "Idle"
    [progressbar_rng] = 0.05
    Application.statusbar = False
    MsgBox ("Error, analysis could not be enabled automatically... please do so manually")
Else: Exit Sub
End If
End Sub
Public Sub start_form(gc As Boolean)
Dim frm As New startform
If gc = True Then
    'close and clear memory
    Unload frm
    Set frm = Nothing
Else
    'waits for the from procedures to be completed
    frm.Show
End If
End Sub
Public Sub month_settings_display()
Dim form As New monthuserform
'the following waits for the from procedures to be completed
form.Show

'close and clear memory
Unload form
Set form = Nothing
End Sub
