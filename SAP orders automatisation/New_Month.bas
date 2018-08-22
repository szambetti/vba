Attribute VB_Name = "New_Month"
Private Sub new_month_base()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'copy cutoff date
    sheets("control panel").range("AA8").copy
    sheets("control panel").range("AA10").PasteSpecial xlPasteValues
    
    'list of columns to be copied
    
    'deprecated array
    Dim column_array_source
    column_array_source = Array("C", "Z", "AR", "BJ", "CB", "CT", "DL", "ED", "EV")
    
    Dim column_array
    column_array = Array("G", "AD", "AV", "BN", "CF", "CX", "DP", "EH", "EZ")
    
    'list of sheets from where to copy tables
    Dim array_sheets
    array_sheets = Array("Daily Orders_3P_QTD", "Daily Orders_QTD", "Daily Orders_3P_YTD", "Daily Orders_YTD")
    
    'loop to copy tables in above list of sheets
    For i = LBound(array_sheets) To UBound(array_sheets)
        
                Worksheets(array_sheets(i)).range("B20:EA242").copy
                Worksheets(array_sheets(i)).range("B270").PasteSpecial xlPasteValues
    
    Next i
    
    'sets "copied with macro" table in DTD orders sheet to 0 in order to have fresh new data for the new month
    Worksheets("Daily Orders_3P_DTD").range("C238:F460") = "0"
    Worksheets("Daily Orders_3P_DTD").range("AB238:AE460") = "0"
    Worksheets("Daily Orders_3P_DTD").range("AO238:AR460") = "0"
    Worksheets("Daily Orders_3P_DTD").range("BB238:BE460") = "0"
    Worksheets("Daily Orders_3P_DTD").range("BO238:BR460") = "0"
    Worksheets("Daily Orders_3P_DTD").range("CB238:CE460") = "0"
    Worksheets("Daily Orders_3P_DTD").range("CO238:CR460") = "0"
    Worksheets("Daily Orders_3P_DTD").range("DB238:DE460") = "0"
    Worksheets("Daily Orders_3P_DTD").range("DO238:DR460") = "0"
    
    
    'since it's new quarter, we don't need old quarter data. only ytd is copied
    Dim array_sheets_2
    array_sheets_2 = Array("Daily Orders_3P_YTD", "Daily Orders_YTD")
    
    For i = LBound(array_sheets_2) To UBound(array_sheets_2)
    
        For j = LBound(column_array) To UBound(column_array)
        
                Worksheets(array_sheets_2(i)).range(column_array_source(j) & "270:" & column_array_source(j) & "492").copy
                Worksheets(array_sheets_2(i)).range(column_array(j) & "20").PasteSpecial xlPasteValues
    
        Next j
    
    Next i
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub
Private Sub erase_data()

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
Private Sub firstdayofthemonth_samequarter()
    
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
    
    MsgBox "Done !" & vbNewLine & vbNewLine & "Now check the periods in the control panel."

End Sub
Private Sub firstdayofthemonth_newquarter()

Call new_month_base

MsgBox "Done !" & vbNewLine & vbNewLine & "Now check the periods in the control panel."

End Sub
