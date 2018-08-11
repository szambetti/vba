'(declarations)
Public custom_cutoff, cutoff As Range, ctrl As Worksheet, xl As Excel.Application, wb As Workbook
Private pt2, pt As PivotTable, progressbar, main_loop, notcorrect As String, lresult As Long
Dim state, inputsave As Variant
' x += 1
Public Function Inc(ByRef x As Range) As Integer
   x.Value = x.Value + 1
End Function
Sub main()

'declare above
Set xl = Excel.Application
Set wb = ThisWorkbook
Set ctrl = ThisWorkbook.Worksheets("control panel")
Set custom_cutoff = ctrl.[custom_cutoff]
Set cutoff = ctrl.[cutoff]
Set pt = wb.sheets("Pivot_Daily Orders").PivotTables("BigPivot")
Set pt2 = wb.sheets("Pivot_Daily Orders").PivotTables("SmallPivot")
state = Array("Running for x = ", "Updating ATLAS", _
            "Refreshing Filters", "Changing pivots", "Finished")
main_loop = Array(1, 2)


'####################
'#   MACRO BEGIN    #
'####################

xl.ScreenUpdating = False

'if a custom cutoff has been speficied cutoff = custom_cutoff
  If custom_cutoff = "" Then
      cutoff = 3
  Else
      cutoff = custom_cutoff.Value
  End If
  ' main loop procedure
  For i = LBound(main_loop) To UBound(main_loop)
  'This the progressbar/cutoff process.
    If main_loop(i) = 1 Then
        With ctrl
            .Activate
            Inc cutoff
            .[cutoff].Value = cutoff
            .Range("AC32") = state(i) & cutoff
            'copies only cutoff date
            .[today_x].copy
            .[today_cp].PasteSpecial xlPasteValues
          End With
        With xl
            .CutCopyMode = False
            .StatusBar = "Running. Please stay idle..."
            DoEvents
        End With
        'update SAP data maybe impplement call instead of variable to free up memory
        Call Application.Run("SAPExecuteCommand", "RefreshData", "ALL")
        DoEvents
    ElseIf main_loop(i) = 2 Then
        'need to cancel previous inc cutoff
        If custom_cutoff = "" Then
            cutoff = 3
        Else
            cutoff = custom_cutoff.Value
        End If
        xl.ScreenUpdating = True
        With ctrl
          .Activate
          .[cutoff].Value = cutoff
          .Range("AC32") = state(UBound(state)) & cutoff
        End With
        xl.ScreenUpdating = False
        Call copy
    Else
        MsgBox "Error on procedure " & main_loop(i)
    End If

    'call SAP filter refresh sub
    Call loop_filters

    'refresh pivots'
    pt.RefreshTable
    pt2.RefreshTable
    DoEvents

  Next i

With xl
    .ScreenUpdating = True
    .StatusBar = False
End With

Saving:

  inputsave = MsgBox("Would you like to save on ShareDrive," & vbNewLine & "SharePoint EPMS and on your desktop this report?", vbYesNo, "DAILY ORDERS")
  If inputsave = vbYes Then
      Call file_save
      DoEvents
      GoTo OpenChrome
  ElseIf inputsave = vbNo Then End If

Exit Sub

OpenChrome:

Dim x As Variant
Dim Path As String

Path = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

x = Shell(Path, vbNormalFocus)
End Sub
Sub loop_filters()

'code for this sub forked from Reagan MacDonald. https://blogs.sap.com/2017/02/03/analysis-for-office-variables-and-filters-via-vba/

' Set up some variables first
Dim result, mytable As ListObject
Dim paramatersarray As Variant
Dim Field_Value As String, Filter_Value As String, Field_Datasource As String, Field_Type As String
Dim myloop, mainloop As Integer

' Lets put the PARAMETERS table into memory in an internal array
' We only want the data, not the headings, so we refer to DataBodyRange on the table
' Now take a copy of the parameters table into memory

Set mytable = ctrl.ListObjects("Parameters")
paramatersarray = mytable.DataBodyRange

'Now lets loop through the parameters array.
' LBound = Lower Bound, the lowest record number (in this case 1)
' UBound = Upper Bound, the highest record number

myloop = 0 ' setting this to something that won't match during the first loop, as we know everything starts from 1 in the parameter file

For mainloop = LBound(paramatersarray) To UBound(paramatersarray)

   ' put the fields into easier variables for the moment
   Field_Loopnum = paramatersarray(mainloop, 1)
   Field_Datasource = paramatersarray(mainloop, 2)
   Field_Type = paramatersarray(mainloop, 3)
   Field_Field = paramatersarray(mainloop, 4)
   Field_Value = paramatersarray(mainloop, 5)

   ' We also want to see if there is a 'next' record, to determine if we are on the last record of the current loop.
   If (mainloop + 1) > UBound(paramatersarray) Then
       Field_Loopnum_next = "No more records" ' Doesn't matter what this is, as long as its not the number of the last loop entry
   Else
       Field_Loopnum_next = paramatersarray(mainloop + 1, 1)
   End If

   If Field_Loopnum <> myloop Then
   ' get ready for a new set of variables
         ' We are going to process variables first so lets turn off variable submissions to keep the speed up

       Call Application.Run("SAPSetRefreshBehaviour", "Off")
       Call Application.Run("SAPExecuteCommand", "PauseVariableSubmit", "On")

   End If


   ' If we find a VARIABLE field, action it.
   If Field_Type = "VARIABLE" Then

       result = Application.Run("SAPSetVariable", Field_Field, Field_Value, "INPUT_STRING", Field_Datasource)

   End If


   If Field_Loopnum <> Field_Loopnum_next Then
   ' the next row in our parameters table has a different loop number, so lets unpause things to refresh the data

       Call Application.Run("SAPExecuteCommand", "PauseVariableSubmit", "Off")

       ' At this stage we have set up refreshed the data for variables, but the user may want to then set some filters.
       ' Loop through the same parameters array from the start again, but this time for the same loop number as the variables and look for filters this time.

       For i = LBound(paramatersarray) To UBound(paramatersarray)

           Filter_Loopnum = paramatersarray(i, 1)
           Filter_DataSource = paramatersarray(mainloop, 2)
           Filter_Type = paramatersarray(mainloop, 3)
           Filter_Field = paramatersarray(mainloop, 4)
           Filter_Value = paramatersarray(mainloop, 5)

           If Filter_Loopnum = Field_Loopnum And Filter_Type = "FILTER" Then

               result = Application.Run("SAPSetFilter", Filter_DataSource, Filter_Field, Filter_Value, "INPUT_STRING")

           End If

       Next i

       ' We are all done with the filters for this specific loop Now refresh whats on screen

       Call Application.Run("SAPSetRefreshBehaviour", "On")

   End If

   ' make a note of the loop we just finished as we will check it in the next loop

   myloop = Field_Loopnum

Next mainloop


End Sub

Sub copy()

    'copy mtd table from yesterday to dtd
    sheets("Daily Orders_3P_MTD").Range("B20:EA242").copy
    sheets("Daily Orders_3P_DTD").Range("B238").PasteSpecial xlPasteValues

    'copy cutoff date
    With ctrl
        .[today_x].copy
        .[today_cp].PasteSpecial xlPasteValues
    End With

    xl.CutCopyMode = False

End Sub
Sub file_save()

    'still using deprecated code

    xl.DisplayAlerts = False
    xl.ScreenUpdating = False

    'Save this workbook on ShareDrive .xlsb
    wb.SaveAs Filename:=Left(ThisWorkbook.Path, 77) _
    & "\" & ctrl.Range("AA22") _
    & "\" & ctrl.Range("AA21") _
    & "\" & ctrl.Range("AA20"), _
    FileFormat:=xlExcel12, CreateBackup:=False

    'State sheet list (array) to be hidden before saving. If sheets names have been changed, just update within the array below
    Dim sheet_list
    sheet_list = Array( _
    "Recon_ATLAS Supply_Weekly", _
    "Recon_ATLAS Demand_Weekly", _
    "RepUnits missing_Weekly", _
    "Pivot_Daily Orders Supply", _
    "Pivot_Daily Orders", _
    "ATLAS_Data", _
    "ATLAS notassig Demand Coun", _
    "Days 2018", _
    "Instructions", _
    "control panel" _
    )

    'Get path to user's desktop - this is done through powershell
    Dim Path As String
    desktop_path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"

    'Store this workbookname in variable wb_name // UNUSED
    Dim wb_name As String
    wb_name = wb.Name

    'Loop for hiding sheets, usually this should remain unchanged
    For i = LBound(sheet_list) To UBound(sheet_list)
        sheets(sheet_list(i)).Visible = False
    Next i

    Worksheets("Daily_Tables").Range("A1").Select
    Columns("M").EntireColumn.Hidden = True

     'Saves on sharepoint GM&S .xlsx
    Workbooks(wb.Name).SaveAs Filename:= _
        "\\sites.abb.com\sites\EPMarketingandSales\Finance\Global MS\" & _
        ctrl.Range("AA22") & "\Daily Demand Orders\" & _
        ctrl.Range("AA21") _
        & "\" & ctrl.Range("AA19"), _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

    'Save on users's desktop .xlsx
    Workbooks(wb.Name).SaveAs Filename:=desktop_path & ctrl.Range("AA19") _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "This is the version saved on your desktop." & vbNewLine & _
    "I've also saved the original on the ShareDrive and a copy on the SharePoint." _
    & vbNewLine & " " & vbNewLine & "You can now send by mail the file on you desktop."
End Sub
