Call generate

Private sub generate()
	dim fso: set fso = CreateObject("Scripting.FileSystemObject")
	dim exl
	dim xlBook
	dim dir
	dim msg
	dim filename
	filename = "template.xlsb"
	dir = fso.GetAbsolutePathName(".")
	set exl = CreateObject("Excel.Application")
	exl.Visible = True

	' a sort of try catch block to look if filename exists
	If (fso.FileExists(dir & "\" & filename)) Then
		' Open template workbook
		Set xlBook = exl.Workbooks.Open(dir & "\" & filename,, FALSE)
	Else
			msg = "Error: to create the report the template file must be named " & filename & " and be placed in the same directory as 'create report.vbs'"
		Exit Sub
	End if

	'runs main macro contained in template.xlsb
	exl.run "main"

	msg = MsgBox ("Macro finished running. Would you like to close this window and exit?", vbyesno)

	'if msgbox is yes then exit
	if msg = 6 then
		exl.close
		Exit Sub
	else
		'do nothing
	end if

	Set xlBook = Nothing
	Set exl = Nothing
	msg = ""
End Sub