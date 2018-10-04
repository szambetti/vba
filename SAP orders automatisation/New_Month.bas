Attribute VB_Name = "New_Month"
'THIS MODULE IS USING SOME OLD, DEPRECATED in parts of the macro
'NEEDS TO BE UPDATED

Sub new_month_base()

 Dim newmonth_table As ListObject
 Dim montharray As Variant, i, j, k As Integer
 Dim ws, ctrl As Worksheet, rng As String

 Set ctrl = Worksheets("control panel")

 Application.DisplayAlerts = False
 Application.ScreenUpdating = False

 'copy cutoff date
 [today_x].copy
 [today_pasted].PasteSpecial xlPasteValues

 Set newmonth_table = ctrl.ListObjects("NewMonth")
 Let montharray = newmonth_table.DataBodyRange

 'loop to copy cells from all sheets, coming from column 2 of ctrl panel month table
 For i = LBound(montharray) To UBound(montharray)
     If montharray(i, 2) <> "" Then
        Set ws = Worksheets(montharray(i, 2))
        With ws
                .range("B20:EA242").copy
                .range("B270").PasteSpecial xlPasteValues
            End With
     End If
 Next i

 'sets "copied with macro" table in DTD orders sheet to 0 in order to have fresh new data for the new month
 [DTD_rng] = "0"

 Let i = 0
 'since it's new quarter, we don't need old quarter data. only ytd is copied
 For j = LBound(montharray) To UBound(montharray)
    If montharray(j, 1) <> "" Then
        Set ws = Worksheets(montharray(j, 1))
        With ws
            For i = LBound(montharray) To UBound(montharray)

               Let rng_source = montharray(i, 3)
               Let rng_paste = montharray(i, 4)

               If (rng_paste <> "" And rng_source <> "") Then

                   .range(rng_source & "270:" & rng_source & "492").copy
                   .range(rng_paste & "20").PasteSpecial xlPasteValues

               End If
            Next i
        End With
    End If
Next j

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub
Public Sub erase_data()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim sheets
    array_sheets = Array("Daily Orders_3P_QTD", "Daily Orders_QTD", "Daily Orders_3P_YTD", "Daily Orders_YTD")

    Dim column_array
    column_array = Array("G", "AD", "AV", "BN", "CF", "CX", "DP", "EH", "EZ")

    For i = LBound(array_sheets) To UBound(array_sheets)

        For j = LBound(column_array) To UBound(column_array)

                Worksheets(array_sheets(i)).range(column_array(j) & "20:" & column_array(j) & "242") = "0"

        Next j

    Next i

    ThisWorkbook.Connections(Worksheets("control panel").range("AF10") & "_demand").Refresh

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Data erased and abacus demand connections refreshed."

End Sub
Public Sub firstdayofthemonth_samequarter()

    Call new_month_base

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    'named arrays differently as the may conflict with macro "new quarter"
    Dim array_sheets_qtd
    array_sheets_qtd = Array("Daily Orders_3P_QTD", "Daily Orders_QTD")

    Dim column_array_source
    column_array_source = Array("C", "Z", "AR", "BJ", "CB", "CT", "DL", "ED", "EV")

    Dim column_array
    column_array = Array("G", "AD", "AV", "BN", "CF", "CX", "DP", "EH", "EZ")

    For i = LBound(array_sheets_qtd) To UBound(array_sheets_qtd)

        For j = LBound(column_array) To UBound(column_array)

                Worksheets(array_sheets_qtd(i)).range(column_array_source(j) & "270:" & column_array_source(j) & "492").copy
                Worksheets(array_sheets_qtd(i)).range(column_array(j) & "20").PasteSpecial xlPasteValues

        Next j

    Next i

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Done !"

End Sub
Public Sub firstdayofthemonth_newquarter()

Call new_month_base

MsgBox "Done !"

End Sub
