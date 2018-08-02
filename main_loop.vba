Sub AO_update()

' Set up some variables first
Dim lResult, mytable As ListObject
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
   
       lResult = Application.Run("SAPSetVariable", Field_Field, Field_Value, "INPUT_STRING", Field_Datasource)

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

               lResult = Application.Run("SAPSetFilter", Filter_DataSource, Filter_Field, Filter_Value, "INPUT_STRING")

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
