Attribute VB_Name = "main__"
'declarations

'global variables
Public uinput As Integer

'local variables
Private custom_cutoff, cutoff, statusbar_rng, progress_rng As range, ctrl, tables As Worksheet, xl As Excel.Application, wb As Workbook
Private pt2, pt As PivotTable, progressbar, main_loop, input_cutoff, k As Integer, notcorrect As String, lResult As Long
Dim state, inputsave As Variant
'update progressbar
Public Sub infostatus(ByRef k As Integer)
    xl.ScreenUpdating = True
    Inc k
    With tables
        .Activate
        .[tables_A1].Select
    End With
    Let progressbar = k
    Let progress_rng.Value = progressbar
    statusbar_rng.Value = state(progressbar - 1)
    DoEvents
    DoEvents
    DoEvents
    xl.ScreenUpdating = False
End Sub
Public Sub Main()

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
Call freeze(True)
Call infostatus(k)

' get custom cutoff from user
InputBox:
Let input_cutoff = InputBox("Hello! Would you like to run the report for how many days ago?" & _
vbNewLine & vbNewLine & "Please insert cutoff (defaults to 3):", "Daily orders generator", 3)
If (input_cutoff = "" Or input_cutoff = 0) Then
    input_cutoff = 0
    inputquestion = MsgBox("Warning!" & vbNewLine & "You selected a cutoff of 0 days." & vbNewLine & _
    "The report will be run for today and data may not be available yet." _
    & vbNewLine & "Is this correct? (Cancel to exit generation)", vbYesNoCancel, "?")
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
            .[today_x_copy].copy
            .[saved_month].PasteSpecial xlPasteValues
          End With
        xl.CutCopyMode = False
        DoEvents
        Call infostatus(k)
        'update all SAP DBs
        Call Application.Run("SAPExecuteCommand", "RefreshData", "ALL")
        DoEvents
    ElseIf main_loop(i) = 2 Then
        'need to cancel previous inc cutoff
        If custom_cutoff = 3 Then
            cutoff = 3
        Else
            cutoff = custom_cutoff.Value
        End If
        ctrl.[cutoff].Value = cutoff
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
    
    'checks if mistake is present in atlas
    If (xl.WorksheetFunction.IsNA(tables.[total_allmarkets_mtd]) = True Or _
    xl.WorksheetFunction.IsError(tables.[total_allmarkets_mtd]) = True) Then
        GoTo NAError
    End If

  Next i
  
  Call freeze(False)
  
  'weekly orders prompt
  weekly_msg = MsgBox("Would you like to run a weekly orders update?" & vbNewLine & _
  "(Do NOT run if an update is not out)", vbYesNo)
    If weekly_msg = vbYes Then
        weekly_msg_confirm = MsgBox("Are you sure?" & vbNewLine & _
        "N.B: if an update is not available (or already downloaded) and you proceed, " & _
        "you WILL break the file.", vbYesNo)
            If weekly_msg_confirm = vbYes Then
                Call weekly
            Else: GoTo Saving
            End If
    Else: GoTo Saving
    End If

  Call infostatus(k)

' goto references

Saving:
    inputsave = MsgBox("Report generated." & vbNewLine & vbNewLine & _
    "Would you like to save on ShareDrive," & vbNewLine & "SharePoint EPMS and on your desktop this report?", _
    vbYesNo, "Daily orders generator")
        Select Case inputsave
            Case vbYes
                Call file_save
                Call infostatus(5)
                GoTo OpenChrome
                Exit Sub
            Case Else
                Call infostatus(5)
        End Select
    Exit Sub

OpenChrome:
    Dim open_chrome As Variant
    Dim Path As String
    Path = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    open_chrome = Shell(Path, vbNormalFocus)
    Call freeze(False)
    Exit Sub

ErrInput:
    MsgBox ("You did not select a cutoff and the macro stopped.")
    Exit Sub

NAError:
     MsgBox ("Looks like there is some new reporting unit present in ATLAS" & _
     "(or even an unknown mistake in ATLAS source worksheets)." & vbNewLine & _
     vbNewLine & "Please fix it in the control panel or ATLAS_Daily, save the template and rerun the launcher.")
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

' Now lets loop through the parameters array.
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
    [month_table].copy
    [month_table_paste].PasteSpecial xlPasteValues
    
    'copy cutoff date
    With ctrl
        .[today_x_copy].copy
        .[saved_month].PasteSpecial xlPasteValues
        .[today_x].copy
        .[today_pasted].PasteSpecial xlPasteValues
    End With
    
    xl.CutCopyMode = False

End Sub
Public Sub file_save()
        
    'State sheet list (array) to be hidden before saving. If sheets names have been changed, just update within the array below
    Dim sheet_list As range
    Set sheet_list = [hide_sheets_list]
    Dim sheet_list_val As Variant
    sheet_list_val = sheet_list.Value
    
    [tables_A1].Select
    DoEvents
    
    Call freeze(True, "Saving daily report...")
    
    With ThisWorkbook
    'save template
    .Save
    
    'Save this workbook on ShareDrive .xlsb
    .SaveAs Filename:=[shared_path_name], _
    FileFormat:=50, CreateBackup:=False
    
    'Get path to user's desktop - this is done through powershell
    Dim Path As String
    desktop_path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"

    'Loop for hiding sheets, usually this should remain unchanged
    For Each cl In sheet_list_val
        If cl <> "" Then
            sheets(cl).Visible = False
        End If
    Next cl
    
    [tables_A1].Select
    Columns("M").EntireColumn.Hidden = True
    Columns("Q").EntireColumn.Hidden = True
    
         'Saves on sharepoint GM&S .xlsx
        .SaveAs Filename:=[GMS_SP], _
            FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        
        'Save on users's desktop .xlsx
        .SaveAs Filename:=desktop_path & Worksheets("control panel").range("AA19") _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    End With
    
    
    tables.Activate
    Call freeze(False)
    
    MsgBox "Saving completed.. the file now open is the one on your desktop"
End Sub
Public Sub weekly()

Call freeze(True)

Dim weekly_rng As range
Set weekly_rng = [latest_weekly]

If Right(weekly_rng.Value, 1) = 1 Then
    weekly_rng = Left(weekly_rng.Value, 6) & 2
    [weekly_period] = [last_month].Value & " A"
    [weekly_period_item] = "Orders received 3rd part."
    
    ElseIf Right(weekly_rng.Value, 1) = 2 Then
        weekly_rng = Left(weekly_rng.Value, 6) & 3
        
    ElseIf Right(weekly_rng.Value, 1) = 3 Then
        weekly_rng = Left(weekly_rng.Value, 6) & 4
        
    ElseIf Right(weekly_rng.Value, 1) = "-" Then
        weekly_rng = [this_month].Value & " W1"
        [weekly_period] = [last_month].Value & " W4"
        [weekly_period_item] = "Orders received gross 3rd part."
    Else
        MsgBox "An error occured in the weekly update macro, could not parse period"
End If

'using async update for query data
ThisWorkbook.Connections([saved_year].Value & "_weekly").Refresh
Application.CalculateUntilAsyncQueriesDone


[tables_A1].Select

Call freeze(False)
MsgBox "Weekly orders refreshed"
End Sub


