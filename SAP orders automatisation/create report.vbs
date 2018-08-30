Call generate

Public sub generate()
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
		exl.Workbooks.Open(dir & "\" & filename,, FALSE)
	Else
			msg = "Error: to create the report the template file must be named " & filename & " and be placed in the same directory as 'create report.vbs'"
		Exit Sub
	End if

	'runs main macro contained in template.xlsb
	exl.run "main"

	Set xlBook = Nothing
	Set exl = Nothing
	msg = ""
End Sub
