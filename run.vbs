' Setup the global variables
Dim sl
Dim xlBook

' Launch Excel
set sl = createobject("Excel.Application")

' Make it visible otherwise it doesn’t work
sl.Application.Visible = True
sl.DisplayAlerts = False

' now open the file you want to refresh
Set xlBook = sl.Workbooks.Open("", 0, False)

' Run the Refresh macro contained within the file.
sl.Application.run "main"

' Save the file and close it
xlBook.save
sl.ActiveWindow.close True

' Close Excel
sl.Quit

'Clear out the memory
Set xlBook = Nothing
Set sl = Nothing
