Dim fso: set fso = CreateObject("Scripting.FileSystemObject")


'str_path = WScript.Arguments(0)                           ' This is the folder path, not the file path
str_path = "D:\Truecoverage\EnrolmentReport\31_Jul"
DataDirectory = fso.GetAbsolutePathName(str_path)
Set folder = fso.GetFolder(DataDirectory)

For each file In folder.Files

If fso.GetExtensionName(file) = "xlsx" Then               ' Getting all files in that folder path, where the extension is ".xlsx"

pathOut = fso.BuildPath(DataDirectory, Replace(fso.GetBaseName(file),"XLSX","CSV") + ".csv")
    Dim oExcel
    Set oExcel = CreateObject("Excel.Application")
    Dim oBook
    Set oBook = oExcel.Workbooks.Open(file)
    oBook.SaveAs pathOut, 6
    oBook.Close False
    oExcel.Quit
    'fso.DeleteFile(file)
 End If
 Next