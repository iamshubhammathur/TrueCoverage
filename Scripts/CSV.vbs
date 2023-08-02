' Get the file path from the command line argument
strFilePath = WScript.Arguments.Item(0)


' Create an Excel object
Dim objExcel
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

' Open the Excel workbook
On Error Resume Next
Dim objWorkbook
Set objWorkbook = objExcel.Workbooks.Open(strFilePath)

If Err.Number <> 0 Then
	WScript.Echo "Error while opening: " & Err.Number & " " & Err.Description
	objWorkbook.Close False ' Close your workbook.
   	objExcel.Quit ' Quit the excel program. 
   	WScript.Quit 
   	Err.Clear
End If

' Convert the Excel file to CSV format
objWorkbook.SaveAs Replace(strFilePath, ".xlsx", ".csv"), 6

' Close the workbook and quit Excel
objWorkbook.Close False
objExcel.Quit

' Release the objects from memory
Set objWorkbook = Nothing
Set objExcel = Nothing

' Inform the user that the conversion is complete
WScript.Echo "The file has been converted from XLSX to CSV format."
