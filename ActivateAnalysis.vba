'As soon as the workbook is opened, it checks for the Analysis COMAddIn.
'If it finds Analysis and thinks it isn’t enabled, it enables it.
'If it finds Analysis and thinks its already enabled, it disables it first then re-enables it to make doubly sure its going to function ok.'

Private Sub Workbook_Open()

Dim lResult As Long

    Dim addin As COMAddIn

    For Each addin In Application.COMAddIns

        If addin.progID = "SapExcelAddIn" Then

            If addin.Connect = False Then
                addin.Connect = True
            ElseIf addin.Connect = True Then
                addin.Connect = False
                addin.Connect = True
            End If

        End If

    Next

End Sub
