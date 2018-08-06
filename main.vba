Sub main()

'state main variables
Dim custom_cutoff As Range
Dim cutoff As Variant
Dim xl As Object
Set xl = Excel.Application
Dim pt2 As PivotTable
Dim pt As PivotTable
Dim progressbar As Integer
Dim wb As Object
Set wb = ThisWorkbook
Dim lresult1 As Long
Dim notcorrect As String
notcorrect = ""

'declare macro string states
Dim state(5) As String
state(1) = "Running for x = "
state(2) = "Updating ATLAS"
state(3) = "Refreshing Filters"
state(4) = "Changing pivots"
state(5) = "Finished"

Dim ctrl As Object
Set ctrl = wb.sheets("control panel")

'declare objects such as cutoff
Set custom_cutoff = ctrl.Range("AF30")
Set cutoff = ctrl.Range("AF6")
Set pt = wb.sheets("Pivot_Daily Orders").PivotTables("BigPivot")
Set pt2 = wb.sheets("Pivot_Daily Orders").PivotTables("SmallPivot")

Application.ScreenUpdating = False
Application.EnableEvents = False

'set progressbar values
progressbar = 0
ctrl.Activate
Range("AA33").Value = progressbar
Application.ScreenUpdating = True
DoEvents
Application.ScreenUpdating = False

  If custom_cutoff = "" Then
    cutoff = 3
  Else
    cutoff = custom_cutoff.Value
  End If

    'This the progressbar/cutoff process. It changes the cutoff date as well as Updating
    'the progress bar. Firstly, AO update is run on cutoff + 1 to always get previous
    'day data needed in the DTD sheet. It will appear many times in this macro
    progressbar = progressbar + 1
    With ctrl
      .Activate
      .Range("AF6").Value = cutoff + 1
      .Range("AC32") = state(1) & cutoff + 1
      .Range("AA33").Value = progressbar
    End With
    xl.StatusBar = "Running. Please stay idle..."
    Application.ScreenUpdating = True
    DoEvents
    Application.ScreenUpdating = False
    'end of process'

    'copies only cutoff date'
    sheets("control panel").Range("AA8").copy
    sheets("control panel").Range("AA10").PasteSpecial xlPasteValues
    Application.CutCopyMode = False

    'progressbar process'
    progressbar = progressbar + 1
    With ctrl
      .Range("AC32") = state(2)
      .Range("AA33").Value = progressbar
    End With
    Application.ScreenUpdating = True
    DoEvents
    Application.ScreenUpdating = False

    'update SAP data'
    lresult1 = Application.Run("SAPExecuteCommand", "RefreshData", "ALL")
    DoEvents

    'progressbar'
    progressbar = progressbar + 1
    With ctrl
      .Activate
      .Range("AC32") = state(3)
      .Range("AA33").Value = progressbar
    End With
    Application.ScreenUpdating = True
    DoEvents
    Application.ScreenUpdating = False
    
    'refresh filters using AO update sub
    Call AO_update
    DoEvents
    
    progressbar = progressbar + 1
    With ctrl
    .Activate
      .Range("AC32") = state(4)
      .Range("AA33").Value = progressbar
    End With
    Application.ScreenUpdating = True
    DoEvents
    Application.ScreenUpdating = False

    'refresh pivots'
    pt.RefreshTable
    pt2.RefreshTable
    DoEvents
    '
    'Second process for true cutoff data. Procedures same as above
    '
    progressbar = progressbar + 2
    With ctrl
      .Activate
      .Range("AF6").Value = cutoff
      .Range("AC32") = state(1) & cutoff
      .Range("AA33").Value = progressbar
    End With
    Application.ScreenUpdating = True
    DoEvents
    Application.ScreenUpdating = False
    
    Call copy
    
    progressbar = progressbar + 1
    With ctrl
      .Activate
      .Range("AC32") = state(3)
      .Range("AA33").Value = progressbar
    End With
    Application.ScreenUpdating = True
    DoEvents
    Application.ScreenUpdating = False

    Call AO_update
    DoEvents
    
    progressbar = progressbar + 1
    With ctrl
      .Activate
      .Range("AC32") = state(3)
      .Range("AA33").Value = progressbar
    End With
    Application.ScreenUpdating = True
    DoEvents
    Application.ScreenUpdating = False
    
    pt.RefreshTable
    pt2.RefreshTable
    DoEvents

    progressbar = 10
    With ctrl
      .Activate
      .Range("AC32") = state(5)
      .Range("AA33") = progressbar
    End With
    
    Application.ScreenUpdating = True
    DoEvents
    xl.StatusBar = False
    Application.EnableEvents = True
    
    
Saving:
  inputsave = InputBox(notcorrect & "Would you like to save? (Y/N)")
  
  If inputsave = "y" Then
      Call file_save
      DoEvents
      GoTo OpenChrome
  ElseIf inputsave = "Y" Then
      Call file_save
      DoEvents
      GoTo OpenChrome
  ElseIf inputsave = "n" Then
      'do nothing'
  ElseIf inputsave = "N" Then
      'do nothing'
  Else
  notcorrect = "I did not recognise your input.."
  GoTo Saving
  End If

Exit Sub

OpenChrome:

Dim x As Variant
Dim Path As String

Path = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

x = Shell(Path, vbNormalFocus)



End Sub
Sub AO_update()

' Set up some variables first
Dim lresult, mytable As ListObject
Dim paramatersarray As Variant
Dim Field_Value As String, Filter_Value As String, Field_Datasource As String, Field_Type As String

' Lets put the PARAMETERS table into memory in an internal array
' We only want the data, not the headings, so we refer to DataBodyRange on the table
' Now take a copy of the parameters table into memory

Set mytable = Worksheets("control panel").ListObjects("Parameters")
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
   
       lresult = Application.Run("SAPSetVariable", Field_Field, Field_Value, "INPUT_STRING", Field_Datasource)

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

               lresult = Application.Run("SAPSetFilter", Filter_DataSource, Filter_Field, Filter_Value, "INPUT_STRING")

           End If

       Next i

       ' We are all done with the filters for this specific loop Now refresh whats on screen

       Call Application.Run("SAPSetRefreshBehaviour", "On")

       ' ***************************************************************************
       ' At this point we now have a refreshed file with all the new variables and filters.
       ' Now you can save/saveas/email here before you move on to the next loop.
       ' ***************************************************************************

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
    sheets("control panel").Range("AA8").copy
    sheets("control panel").Range("AA10").PasteSpecial xlPasteValues
    
    Application.CutCopyMode = False

End Sub
Sub file_save()
     
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'Save this workbook on ShareDrive .xlsb
    Workbooks(ThisWorkbook.Name).SaveAs Filename:=Left(ThisWorkbook.Path, 77) _
    & "\" & Worksheets("control panel").Range("AA22") _
    & "\" & Worksheets("control panel").Range("AA21") _
    & "\" & Worksheets("control panel").Range("AA20"), _
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
    wb_name = ThisWorkbook.Name

    'Loop for hiding sheets, usually this should remain unchanged
    For i = LBound(sheet_list) To UBound(sheet_list)
        sheets(sheet_list(i)).Visible = False
    Next i
    
    Worksheets("Daily_Tables").Range("A1").Select
    Columns("M").EntireColumn.Hidden = True
    
     'Saves on sharepoint GM&S .xlsx
    Workbooks(ThisWorkbook.Name).SaveAs Filename:= _
        "\\xxx.xxx\sites\EPMarketingandSales\Finance\Global MS\" & _
        Worksheets("control panel").Range("AA22") & "\Daily Demand Orders\" & _
        Worksheets("control panel").Range("AA21") _
        & "\" & Worksheets("control panel").Range("AA19"), _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    'Save on users's desktop .xlsx
    Workbooks(ThisWorkbook.Name).SaveAs Filename:=desktop_path & Worksheets("control panel").Range("AA19") _
    , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "This is the version saved on your desktop." & vbNewLine & _
    "I've also saved the original on the ShareDrive and a copy on the SharePoint." _
    & vbNewLine & " " & vbNewLine & "You can now send by mail the file on you desktop."
End Sub
