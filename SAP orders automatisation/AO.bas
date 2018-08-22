Attribute VB_Name = "AO"
'(declarations)
Public custom_cutoff, cutoff, statusbar_rng, progress_rng As range, ctrl, tables As Worksheet, xl As Excel.Application, wb As Workbook
Public pt2, pt As PivotTable
Private progressbar, main_loop, input_cutoff, k As Integer, notcorrect As String, lResult As Long
Dim state, inputsave As Variant
'functions
Public Function Inc(ByRef x As Integer) As Integer
   x = x + 1
End Function
Public Sub infostatus(ByRef k As Integer)
    xl.ScreenUpdating = True
    Inc k
    With tables
        .Activate
        .range("A1").Select
    End With
    Let progressbar = k
    Let progress_rng.Value = progressbar
    statusbar_rng.Value = state(progressbar - 1)
    DoEvents
    xl.ScreenUpdating = False
End Sub
Public Sub main()

'declare above
Set xl = Excel.Application
xl.ScreenUpdating = False
Set wb = ThisWorkbook
Set ctrl = ThisWorkbook.Worksheets("control panel")
Set tables = ThisWorkbook.Worksheets("Daily_Tables")
Set custom_cutoff = ctrl.[custom_cutoff]
Set cutoff = ctrl.[cutoff]
Set statusbar_rng = tables.[state_rng]
Set progress_rng = tables.[progressbar_rng]
Set pt = wb.sheets("Pivot_Daily Orders").PivotTables("BigPivot")
Set pt2 = wb.sheets("Pivot_Daily Orders").PivotTables("SmallPivot")
Let state = Array("Running...", "Updating ATLAS", _
            "Refreshing Filters", "Changing pivots", "Saving", "Finished")
Let main_loop = Array(1, 2)
    
'begin user interface
Let k = 0
Call infostatus(k)

' get custom cutoff from user
InputBox:
Let input_cutoff = InputBox("Hello! Would you like to run the report for how many days ago?" & vbNewLine & vbNewLine & "Please insert cutoff (defaults to 3):", "Daily orders generator", 3)
If input_cutoff = "" Then
    input_cutoff = 0
    inputquestion = MsgBox("Warning!" & vbNewLine & "You selected a cutoff of 0 days.. " & vbNewLine & "Is this correct? (Cancel to exit)", vbYesNoCancel)
        Select Case inputquestion
            Case vbYes 'do nothing
            Case vbCancel
                Exit Sub
            Case Else
                GoTo InputBox
        End Select
End If
Let custom_cutoff.Value = input_cutoff

  'if a custom cutoff has been speficied cutoff = custom_cutoff
  If custom_cutoff = 3 Then
      cutoff = 3
  Else
      cutoff = custom_cutoff.Value
  End If
  ' main loop procedure
  For i = LBound(main_loop) To UBound(main_loop)
  'This the progressbar/cutoff process.
    If main_loop(i) = 1 Then
        With ctrl
            cutoff = cutoff + 1
            .[cutoff].Value = cutoff
            'copies only cutoff date and saved on
            .[today_x].copy
            .[today_pasted].PasteSpecial xlPasteValues
            .range("AF8:AF11").copy
            .range("AG8").PasteSpecial xlPasteValues
          End With
        With xl
            .CutCopyMode = False
            .statusbar = "Running. Please stay idle..."
            DoEvents
        End With
        Call infostatus(k)
        'update SAP data maybe impplement call instead of variable to free up memory
        Call Application.Run("SAPExecuteCommand", "RefreshData", "ALL")
        DoEvents
    ElseIf main_loop(i) = 2 Then
        'need to cancel previous inc cutoff
        If custom_cutoff = 3 Then
            cutoff = 3
        Else
            cutoff = custom_cutoff.Value
        End If
        With ctrl
          .[cutoff].Value = cutoff
        End With
        Call copy
    Else
        MsgBox "Error on loop " & main_loop(i)
    End If
    
    'call SAP filter refresh sub
    Call loop_filters
    
    'refresh pivots'
    pt.RefreshTable
    pt2.RefreshTable
    DoEvents
    
    If xl.WorksheetFunction.IsNA(tables.[total_allmarkets_mtd]) = True Or _
    xl.WorksheetFunction.IsError(tables.[total_allmarkets_mtd]) = True Then
        GoTo NAError
    End If

  Next i

  Call infostatus(k)

' goto references

Saving:
    inputsave = MsgBox("Report generated." & vbNewLine & vbNewLine & _
    "Would you like to save on ShareDrive," & vbNewLine & "SharePoint EPMS and on your desktop this report?", _
    vbYesNo, "Daily orders generator")
        Select Case inputsave
            Case vbYes
                With xl
                    .ScreenUpdating = True
                    .statusbar = False
                End With
                Call file_save
                Call infostatus(5)
                DoEvents
                GoTo OpenChrome
                Exit Sub
            Case vbNo
                'do nothing
        End Select
    Call infostatus(5)
    With xl
        .ScreenUpdating = True
        .statusbar = False
    End With
    Exit Sub

OpenChrome:
    Dim open_chrome As Variant
    Dim Path As String
    Path = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    open_chrome = Shell(Path, vbNormalFocus)
    Exit Sub

ErrInput:
    MsgBox ("You did not select a cutoff and the macro stopped.")
    Exit Sub

NAError:
     MsgBox ("Looks like there some new reporting unit reporting in ATLAS (or even an unknown mistake in ATLAS source worksheets)." & vbNewLine & vbNewLine & "Please add it in the control panel, save the template and rerun the macro.")
     Exit Sub
End Sub
Private Sub loop_filters()

'forked from Reagan MacDonald. https://blogs.sap.com/2017/02/03/analysis-for-office-variables-and-filters-via-vba/

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

   End If

   ' make a note of the loop we just finished as we will check it in the next loop

   myloop = Field_Loopnum

Next mainloop

'refresh whats on screen
Call Application.Run("SAPSetRefreshBehaviour", "On")

End Sub
Private Sub copy()

    'copy mtd table from yesterday to dtd
    sheets("Daily Orders_3P_MTD").range("B20:EA242").copy
    sheets("Daily Orders_3P_DTD").range("B238").PasteSpecial xlPasteValues
    
    'copy cutoff date
    With ctrl
        .range("AF8:AF11").copy
        .range("AG8").PasteSpecial xlPasteValues
        .[today_x].copy
        .[today_pasted].PasteSpecial xlPasteValues
    End With
    
    xl.CutCopyMode = False

End Sub
Private Sub file_save()
    
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
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'save template
    ThisWorkbook.Save
    
    'Save this workbook on ShareDrive .xlsb
    Workbooks(ThisWorkbook.Name).SaveAs Filename:= _
    ctrl.range("AA22"), _
    FileFormat:=50, CreateBackup:=False
    
    'Get path to user's desktop - this is done through powershell
    'Dim Path As String
    'desktop_path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"

    'Loop for hiding sheets, usually this should remain unchanged
    For i = LBound(sheet_list) To UBound(sheet_list)
        sheets(sheet_list(i)).Visible = False
    Next i
    
    tables.range("A1").Select
    Columns("M").EntireColumn.Hidden = True
    Columns("Q").EntireColumn.Hidden = True
    
     'Saves on sharepoint GM&S .xlsx
    Workbooks(ThisWorkbook.Name).SaveAs Filename:= _
        ctrl.range("AA23"), _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    'Save on users's desktop .xlsx
    'Workbooks(ThisWorkbook.Name).SaveAs Filename:=desktop_path & Worksheets("control panel").range("AA19") _
    ', FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Saving completed.. the file now open is the file on the M&S Sharepoint"
End Sub
