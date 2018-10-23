Attribute VB_Name = "Module1"
Private Const PPT_EXT = ".pptx"
Private broken As Boolean

' following sub forked from: https://stackoverflow.com/questions/18707249/exporting-powerpoint-sections-into-separate-files
' User: PatricK
' Thanks a lot !
Sub SplitIntoSectionFiles()
    On Error Resume Next
    Dim aNewFiles() As Variant, sPath As String, i As Long, pname, p_name, yymm As String

    'get period
    Let yymm = Format(Date, "yy") & Format(Date, "mm")

    'gets presentation name
    Let pname = ActivePresentation.Name
        If pname Like "MASTER_*" Then
            Let p_name = Mid(pname, 8, InStr(8, pname, ".ppt") - 8)
        ElseIf pname Like "* links broken*" Then
            Let p_name = Left(pname, InStr( _
                1, pname, ".ppt") - 14)
        Else
            Let p_name = Left(pname, InStr( _
                1, pname, ".ppt") - 1)
        End If

    With ActivePresentation
        sPath = .Path & "\"
        For i = 1 To .SectionProperties.Count
            ReDim Preserve aNewFiles(i)
            ' Store the Section Names
            aNewFiles(i - 1) = .SectionProperties.Name(i)
            'create period folder if it does not exist
            If Len(Dir(sPath & yymm, vbDirectory)) = 0 Then
                    MkDir sPath & yymm
            End If
            'save copy into file and open it
            .SaveCopyAs sPath & yymm & "\" & yymm & "_" & p_name & "_" & aNewFiles(i - 1), ppSaveAsOpenXMLPresentation
            ' Call Sub to Remove irrelevant sections
            RemoveOtherSections sPath & yymm & "\" & yymm & "_" & p_name & "_" & aNewFiles(i - 1) & PPT_EXT
            DoEvents
        Next
        If .SectionProperties.Count > 0 And Err.Number = 0 Then
            If broken = True Then
             MsgBox "Successfully split " & .Name & " into " & UBound(aNewFiles) & " files, saved into subfolder '" & yymm & "' without Excel Links."
            Else
             MsgBox "Successfully split " & .Name & " into " & UBound(aNewFiles) & " files, saved into subfolder " & yymm & "."
            End If
        End If
    End With
    Let broken = False

End Sub

Private Sub RemoveOtherSections(sPPT As String)
    On Error Resume Next
    Dim oPPT As Presentation, i As Long

    Set oPPT = Presentations.Open(FileName:=sPPT, WithWindow:=msoFalse)
    With oPPT
        ' Delete Sections from last to first
        For i = .SectionProperties.Count To 1 Step -1
            ' Delete Sections that are not in the file name
            If Not InStr(1, .Name, .SectionProperties.Name(i)) > 0 Then
                ' Delete the Section, along with the slides associated with it
                .SectionProperties.Delete i, True
            End If
        Next
        .Save
        .Close
    End With
    Set oPPT = Nothing
End Sub

'update all excel links macro
Public Sub UpdateLinks()

    On Error Resume Next
    Dim oShape As Shape, oSlide As Slide, uinput As Integer

    uinput = MsgBox("You are about to begin updating all links in this presentation. This will take very long to complete." & vbNewLine & vbNewLine & _
            "Please make sure that all links are in the same format as in the Info tab. Press 'Yes' to continue." & vbNewLine & vbNewLine & _
            "To stop the process whilst the macro is running, press Alt+Break", vbYesNo)
    Select Case uinput
     Case vbYes
      GoTo main
     Case Else: Exit Sub
    End Select

main:
    For Each oSlide In ActivePresentation.Slides
        For Each oShape In oSlide.Shapes
            If oShape.Type = msoLinkedOLEObject Then
                oShape.LinkFormat.Update
            End If
        Next oShape
    Next oSlide
    MsgBox "Done"
End Sub

'break all links in presentation macro
Public Sub BreakLinks()
    Dim oShape As Shape, oSlide As Slide, uinput As Integer

    uinput = MsgBox("You are about to break all links in this presentation and split it into single conutries." & vbNewLine & vbNewLine & _
            "This will take quite long. Press 'Yes' to continue." & vbNewLine & vbNewLine & _
            "To stop the process whilst the macro is running, press Alt+Break", vbYesNo)
    Select Case uinput
     Case vbYes
      GoTo main
     Case Else: Exit Sub
    End Select


main:

    Dim yyyy As String
    Let yyyy = Format(Date, "yyyy")


    With ActivePresentation
        .SaveAs FileName:=ActivePresentation.Path & "\" & yyyy & " links broken.pptm", fileformat:=ppSaveAsOpenXMLPresentationMacroEnabled

        For Each oSlide In .Slides
            For Each oShape In oSlide.Shapes
                If oShape.Type = msoLinkedOLEObject Then
                    oShape.LinkFormat.BreakLink
                End If
            Next oShape
        Next oSlide

        DoEvents
    End With

    Let broken = True
 Call SplitIntoSectionFiles

    With ActivePresentation
        If Len(Dir(.FullName)) Then
            .Saved = True
            SetAttr .FullName, vbNormal
            Dim presname As String
            Let presname = .FullName
            .Close
            Kill presname
        End If
    End With

End Sub
Public Sub help()
MsgBox "1. All links must be in the following form:" & vbNewLine & vbNewLine & "link"
& vbNewLine & vbNewLine & "If a local address is used or an address different from the above, the update will not work" _
& vbNewLine & vbNewLine & "2. Before running macros close all other presentations and excel worksheets. This will help memory and avoid mistakes on events" _
& vbNewLine & vbNewLine & "3. Macros can be edited from the Developer Tab > Visual Basic > Module1" _
& vbNewLine & vbNewLine & "4. For any question ask name"
End Sub
