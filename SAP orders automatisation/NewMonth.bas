Attribute VB_Name = "NewMonth"
'STILL BETA
Public Enum quarter
 NewQuarter = 0
 SameQuarter = 1
End Enum

Public Sub NewMonth_Main(x As quarter)

 Dim newmonth_table As ListObject
 Dim montharray As Variant, i, j, k As Integer
 Dim ws, ctrl As Worksheet, rng As String

 Call freeze(True)
 
 'copy cutoff date
 Set ctrl = Worksheets("control panel")
 ctrl.[today_x].copy
 ctrl.[today_pasted].PasteSpecial xlPasteValues
 
 Set newmonth_table = ctrl.ListObjects("NewMonth")
 Let montharray = newmonth_table.DataBodyRange
 
 'loop to copy cells from all sheets, coming from column 2 of ctrl panel month table
 For i = LBound(montharray) To UBound(montharray)
     If montharray(i, 2) <> "" Then
        Set ws = Worksheets(montharray(i, 2))
        With ws
                .range(montharray(1, 5)).copy
                .range(montharray(2, 5)).PasteSpecial xlPasteValues
            End With
     End If
 Next i
 
 'sets "copied with macro" table in DTD orders sheet to 0 in order to have fresh new data for the new month
 [DTD_rng] = "0"
 
 'loop to copy data in tables below(starting from 270) and copying it in Previous LM columns
 Let i = 0
 'since it's new quarter, we don't need old quarter data. only ytd is copied
 For j = LBound(montharray) To UBound(montharray)
    If montharray(j, 1) <> "" Then
        Set ws = Worksheets(montharray(j, 1))
        With ws
            For i = LBound(montharray) To UBound(montharray)
            
               Dim rng_paste, rng_source As String
               Let rng_source = montharray(i, 3)
               Let rng_paste = montharray(i, 4)
               
               If (rng_paste <> "" And rng_source <> "") Then
               
                   .range(rng_source & montharray(4, 5) & ":" & rng_source & montharray(5, 5)).copy
                   .range(rng_paste & montharray(6, 5)).PasteSpecial xlPasteValues
                   
               End If
            Next i
        End With
    End If
Next j
    
'lastest weekly is set to - to get 0 data

[latest_weekly] = Left([this_month].Value, 4) & " W-"
[weekly_period] = Left([last_month].Value, 4) & " -"
ThisWorkbook.Connections([saved_year].Value & "_weekly").Refresh
Application.CalculateUntilAsyncQueriesDone
    
'Distincts between new and same quarter, as they require different data approaches
If x = NewQuarter Then

    [weekly_last_quarter] = Left([last_month].Value, 4) & " W4"
    [current_quarter_bgt] = Left([this_month].Value, 4) & " B"

Else
    Let i = 0
    Let j = 0
    'since it's new quarter, we don't need old quarter data. only ytd is copied
    For j = LBound(montharray) To UBound(montharray)
        If montharray(j, 2) <> "" Then
            Set ws = Worksheets(montharray(j, 2))
            With ws
                For i = LBound(montharray) To UBound(montharray)
                
                   Let rng_source = montharray(i, 3)
                   Let rng_paste = montharray(i, 4)
                   
                   If (rng_paste <> "" And rng_source <> "") Then
                   
                    .range(rng_source & montharray(4, 5) & ":" & rng_source & montharray(5, 5)).copy
                    .range(rng_paste & montharray(6, 5)).PasteSpecial xlPasteValues
                       
                   End If
                Next i
            End With
        End If
    Next j
End If
    
Call freeze(False)
    
MsgBox "Done !"

End Sub
Public Sub EraseCopiedData()

    'sub used to delete at abacus first release previously copied data at month beginning

    Dim newmonth_table As ListObject
    Dim montharray As Variant, i, j, k As Integer
    Dim ws, ctrl As Worksheet, rng As String
    Set ctrl = Worksheets("control panel")
    Set newmonth_table = ctrl.ListObjects("NewMonth")
    Let montharray = newmonth_table.DataBodyRange
 

    Call freeze(True)
    
    Let i = 0
    Let j = 0
    'loop to copy cells from all sheets, coming from column 2 of ctrl panel month table
    For i = LBound(montharray) To UBound(montharray)
        If montharray(i, 2) <> "" Then
           Set ws = Worksheets(montharray(i, 2))
           With ws
                For j = LBound(montharray) To UBound(montharray)
                    Dim rng_paste As String
                    Let rng_paste = montharray(j, 4)
                    If (rng_paste <> "") Then
                        .range(rng_paste & montharray(6, 5) & ":" & rng_paste & montharray(7, 5)) = "0"
                    End If
                Next j
           End With
        End If
    Next i
    
    ThisWorkbook.Connections([saved_year].Value & "_demand").Refresh
    Application.CalculateUntilAsyncQueriesDone
    
    Call freeze(False)

    MsgBox "Data erased and abacus demand connections refreshed."
    
End Sub
