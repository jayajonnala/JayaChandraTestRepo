
'Option Explicit
'
'Dim ExcelApp, Workbook, Worksheet
'Dim dict, lastRow, i, colIndex, cellValue, dataToCopy
'
'' === CONFIGURATION ===
'Dim excelFilePath
'excelFilePath = "C:\Users\JJONN\Desktop\Jaya_Chandra.XLSX"  ' Change this to your actual file path
'colIndex = 1  ' Column A = 1, B = 2, etc.
'
'' === Create Excel Application ===
'Set ExcelApp = CreateObject("Excel.Application")
'ExcelApp.Visible = False
'Set Workbook = ExcelApp.Workbooks.Open(excelFilePath)
'Set Worksheet = Workbook.Sheets(1)
'
'' === Create Dictionary to Track Unique Values ===
'Set dict = CreateObject("Scripting.Dictionary")
'
'' === Find Last Row in the Column ===
'lastRow = Worksheet.Cells(Worksheet.Rows.Count, colIndex).End(-4162).Row  ' xlUp = -4162
'
'' === Loop Through Rows and Collect Unique Values ===
'dataToCopy = ""
'For i = 2 To lastRow  ' Assuming row 1 is header
'    cellValue = Trim(CStr(Worksheet.Cells(i, colIndex).Value))
'    If Not dict.Exists(cellValue) And cellValue <> "" Then
'        dict.Add cellValue, True
'        dataToCopy = dataToCopy & cellValue & vbCrLf
'    End If
'Next
'
'' === Copy to Clipboard ===
'CopyToClipboard dataToCopy
'
'' === Cleanup ===
'Workbook.Close False
'ExcelApp.Quit
'Set Worksheet = Nothing
'Set Workbook = Nothing
'Set ExcelApp = Nothing
'Set dict = Nothing
'
'MsgBox "Unique values copied to clipboard!", vbInformation
'
'' === Function to Copy Text to Clipboard ===
'Sub CopyToClipboard(text)
'    Dim htmlfile
'    Set htmlfile = CreateObject("htmlfile")
'    htmlfile.ParentWindow.ClipboardData.SetData "Text", text
'End Sub



Option Explicit

Dim ExcelApp, Workbook, Worksheet
Dim dict, lastRow, i, colIndex, cellValue, outputText
Dim fso, outputFile

' === CONFIGURATION ===
Dim excelFilePath, outputFilePath
excelFilePath = "C:\Users\JJONN\Desktop\Jaya_Chandra.XLSX"       ' Ì†ΩÌ¥Å Update this path
outputFilePath = "C:\Users\JJONN\Desktop\Jaya_Chandra.txt" ' Ì†ΩÌ¥Å Output Notepad file path
colIndex = 1  ' Ì†ΩÌ¥Å Column A = 1, B = 2, etc.

' === Create Excel Application ===
Set ExcelApp = CreateObject("Excel.Application")
ExcelApp.Visible = False
Set Workbook = ExcelApp.Workbooks.Open(excelFilePath)
Set Worksheet = Workbook.Sheets(1)

' === Create Dictionary to Track Unique Values ===
Set dict = CreateObject("Scripting.Dictionary")

' === Find Last Row in the Column ===
lastRow = Worksheet.Cells(Worksheet.Rows.Count, colIndex).End(-4162).Row  ' xlUp = -4162

' === Loop Through Rows and Collect Unique Values ===
outputText = ""
For i = 2 To lastRow  ' Assuming row 1 is header
    cellValue = Trim(CStr(Worksheet.Cells(i, colIndex).Value))
    If cellValue <> "" Then
        If Not dict.Exists(cellValue) Then
            dict.Add cellValue, True
            outputText = outputText & cellValue & vbCrLf
        End If
    End If
Next

' === Write to Notepad File ===
Set fso = CreateObject("Scripting.FileSystemObject")
Set outputFile = fso.CreateTextFile(outputFilePath, True)
outputFile.Write outputText
outputFile.Close

' === Cleanup ===
Workbook.Close False
ExcelApp.Quit
Set Worksheet = Nothing
Set Workbook = Nothing
Set ExcelApp = Nothing
Set dict = Nothing
Set fso = Nothing
Set outputFile = Nothing


