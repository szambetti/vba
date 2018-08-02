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
    Dim lResult As Long
    
    'declare macro string states
    Dim state(5) As String
    state(1) = "Running for x = "
    state(2) = "Updating ATLAS"
    state(3) = "Refreshing Filters"
    state(4) = "Changing pivots"
    state(5) = "Finished"
    
    Dim screenupdate As Boolean
    screenupdate = xl.ScreenUpdating
    Dim ctrl As Object
    Set ctrl = wb.sheets("control panel")
    
    Set custom_cutoff = ctrl.Range("AG6")
    Set cutoff = ctrl.Range("AF6")
    Set pt = wb.sheets("Pivot_Daily Orders").PivotTables("BigPivot")
    Set pt2 = wb.sheets("Pivot_Daily Orders").PivotTables("SmallPivot")

    ctrl.Activate
    progressbar = 0
    Range("A26").Value = progressbar
    
            screenupdate = False
            
            If custom_cutoff = "" Then
            cutoff = 3
            Else
            cutoff = custom_cutoff.Value
            End If
            
              progressbar = progressbar + 1
            
              With ctrl
                  .Range("AF6").Value = cutoff + 1
                  .Range("C25") = state(1) & cutoff + 1
                  .Range("A26").Value = progressbar
              End With
            
              
            xl.StatusBar = "Running. Please stay idle..."
              
            screenupdate = Not (screenupdate)
            DoEvents
            screenupdate = Not (screenupdate)
        Call copy
              
            screenupdate = Not (screenupdate)
            progressbar = progressbar + 1
            
              With ctrl
                  .Range("C25") = state(2)
                  .Range("A26").Value = progressbar
              End With
            
            screenupdate = Not (screenupdate)
              
            lResult = Application.Run("SAPExecuteCommand", "RefreshData", "ALL")
            
            DoEvents
            
            screenupdate = Not (screenupdate)
            progressbar = progressbar + 1
            
              With ctrl
                  .Range("C25") = state(3)
                  .Range("A26").Value = progressbar
              End With
        
            screenupdate = Not (screenupdate)
              
        Call AO_update
              
            DoEvents
              
            screenupdate = Not (screenupdate)
            progressbar = progressbar + 1
            
              With ctrl
                  .Range("C25") = state(4)
                  .Range("A26").Value = progressbar
              End With
              
            screenupdate = Not (screenupdate)
        
            pt.RefreshTable
            pt2.RefreshTable
            DoEvents
            
            ctrl.Activate
            
              
            screenupdate = True
            cutoff = 3
            progressbar = progressbar + 2
            
              With ctrl
                  .Range("AF6").Value = cutoff
                  .Range("C25") = state(1) & cutoff
                  .Range("A26").Value = progressbar
              End With
              
            screenupdate = Not (screenupdate)
              
        Call copy
              
            screenupdate = Not (screenupdate)
            progressbar = progressbar + 1
            
            With ctrl
                  .Range("C25") = state(3)
                  .Range("A26").Value = progressbar
            End With
              
        Call AO_update
              
              DoEvents
              
            screenupdate = Not (screenupdate)
            progressbar = progressbar + 1
            
            With ctrl
                .Activate
                .Range("C25") = state(3)
                .Range("A26").Value = progressbar
            End With
              
            screenupdate = Not (screenupdate)
              
              pt.RefreshTable
              pt2.RefreshTable
              
              DoEvents
              
              ctrl.Activate
              
            screenupdate = Not (screenupdate)
            progressbar = 10
              
              With ctrl
              
              .Range("C25") = state(5)
              .Range("A26") = progressbar
              
              End With
              
            xl.StatusBar = False
              
            Exit Sub
    
OpenChrome:
    
    Dim x As Variant
    Dim Path As String
    
    Path = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

    x = Shell(Path, vbNormalFocus)
      
    

End Sub
