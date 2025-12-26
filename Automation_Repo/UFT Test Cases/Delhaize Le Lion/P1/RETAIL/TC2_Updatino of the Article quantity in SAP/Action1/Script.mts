
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_136_SRM Regularization SC Generic item on cost Center (MI)
'.................Author : TCS 	   :
'................ Creation Date    :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	DataRowSet= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


gstrTestCaseName = "TC1_Retrival of article from the excel"
'gstrTestCaseName = "Test_Manhattan_GR_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\TASEWork\Data\TASE_DT_114_SRM Standard SC Free Text on Cost Cent P1_.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------

Call ExcelFileTransfer(DT_ContractSourceFolderPath,DT_ContractTargetFilePath,DT_ContractTargetFile,21)
Call SaveXmlAsXlsx(DT_ContractTargetFilePath,DT_EXCEL_PATH)
RowNum =  FindFirstRow_Zero_ByPath_NoCleanup(DT_EXCEL_PATH, DT_SHEET_NAME, 9, True)

Call WriteRunTimeDataToExcel ("DT_ROW_NUMBER",RowNum)

DT_SAP_ID =  GetValue_ByRowCol(DT_EXCEL_PATH, DT_SHEET_NAME, DT_ROW_NUMBER, 2)
call WriteRunTimeDataToExcel ("DT_SAP_ID",DT_SAP_ID)
DT_NET_QTY = GetValue_ByRowCol(DT_EXCEL_PATH, DT_SHEET_NAME, DT_ROW_NUMBER, 18)
call WriteRunTimeDataToExcel ("DT_NET_QTY",DT_NET_QTY)

Call FinalStatus()

Public Function GetValue_ByRowCol(excelPath, sheetName, rowNum, colNum)
    On Error Resume Next
strStepName = " Get Value By Row number from excel in column:" &colNum
    Dim fso, xlApp, wb, ws, valueOut
    valueOut = ""

    ' Check file exists
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(excelPath) Then
        Reporter.ReportEvent micFail, "GetValue_ByRowCol", "File not found: " & excelPath
 					
        GetValue_ByRowCol = ""
        						strStatus = "FAIL"
							strMsg = "File not found: """&excelPath&""""
							Call ReporterFunction(strLibraryFileName,"GetValue_ByRowCol","1","File not found: ", strMsg)
        Exit Function
    End If

    ' Open Excel
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    Set wb = xlApp.Workbooks.Open(excelPath)
    Set ws = wb.Worksheets(sheetName)

    If ws Is Nothing Then
        Reporter.ReportEvent micFail, "GetValue_ByRowCol", "Sheet not found: " & sheetName
        GetValue_ByRowCol = ""
        						strStatus = "FAIL"
							strMsg = "Sheet not found: """&excelPath&""""
							Call ReporterFunction(strLibraryFileName,"GetValue_ByRowCol","1","Sheet not found: ", strMsg)
        Exit Function
    End If

    ' Read cell value
    valueOut = ws.Cells(rowNum, colNum).Value
							strStatus = "Pass"
							strMsg = "Value rtrived from the Col & Row: """&valueOut&""""
							Call ReporterFunction(strLibraryFileName,"GetValue_ByRowCol","2","Value retrived from the Excel with Row  & Col  as reference: ", strMsg)
    GetValue_ByRowCol = valueOut
    
    If blnDefault_eSwiftReporting Then  
		   Call UpdateResultHtml(strStepName,valueOut,strMsg,strStatus,"")
	   End If
    xlApp.Quit
    Set wb = Nothing
    Set xlApp = Nothing
    Set fso = Nothing
End Function

Public Function FindFirstRow_Zero_ByPath_NoCleanup(excelPath, sheetName, colIndex, hasHeader)
    On Error Resume Next

    Const xlValues = -4163
    Const xlWhole  = 1
    Const xlByRows = 1
    Const xlNext   = 1
    Const xlUp     = -4162

    Dim fso, xlApp, wb, ws
    Dim startRow, lastRow
    Dim rngCol, foundCell
    Dim resultRow, v

    resultRow = 0
	strStepName = "Finding the Row number with value Zero"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(excelPath) Then
        Reporter.ReportEvent micFail, "FindFirstRow_Zero", "File not found: " & excelPath
        						strStatus = "FAIL"
							strMsg = "File not found: """&excelPath&""""
							Call ReporterFunction(strLibraryFileName,"FindFirstRow_Zero","1","File not found: ", strMsg)
        FindFirstRow_Zero_ByPath_NoCleanup = 0
        Exit Function
    End If

    Set xlApp = CreateObject("Excel.Application")
    If Err.Number <> 0 Or xlApp Is Nothing Then
        Reporter.ReportEvent micFail, "FindFirstRow_Zero", "Could not start Excel: " & Err.Description
       						strStatus = "FAIL"
							strMsg = "Could not start Excel: """&excelPath&""""
							Call ReporterFunction(strLibraryFileName,"FindFirstRow_Zero","1","Could not start Excel: ", strMsg)
        Err.Clear
        FindFirstRow_Zero_ByPath_NoCleanup = 0
        Exit Function
    End If

    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    Set wb = xlApp.Workbooks.Open(excelPath)
    If Err.Number <> 0 Or wb Is Nothing Then
        Reporter.ReportEvent micFail, "FindFirstRow_Zero", "Failed to open workbook: " & Err.Description
       						 strStatus = "FAIL"
							strMsg = "Failed to open workbook: """&excelPath&""""
							Call ReporterFunction(strLibraryFileName,"FindFirstRow_Zero","1","Failed to open workbook: ", strMsg)
        Err.Clear
        FindFirstRow_Zero_ByPath_NoCleanup = 0
        Exit Function
    End If

    Set ws = wb.Worksheets(sheetName)
    If ws Is Nothing Then
        Reporter.ReportEvent micFail, "FindFirstRow_Zero", "Worksheet not found: " & sheetName
         strStatus = "FAIL"
							strMsg = "Worksheet not found: """&excelPath&""""
							Call ReporterFunction(strLibraryFileName,"FindFirstRow_Zero","1","Worksheet not found: ", strMsg)
        FindFirstRow_Zero_ByPath_NoCleanup = 0
        Exit Function
    End If

    If hasHeader Then
        startRow = 2
    Else
        startRow = 1
    End If

    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
    If lastRow < startRow Then
        FindFirstRow_Zero_ByPath_NoCleanup = 0
        Exit Function
    End If

    Set rngCol = ws.Range(ws.Cells(startRow, colIndex), ws.Cells(lastRow, colIndex))

    ' Fast path: positional arguments (VBScript-safe)
    Set foundCell = rngCol.Find(0, rngCol.Cells(rngCol.Rows.Count, rngCol.Columns.Count), _
                                xlValues, xlWhole, xlByRows, xlNext, False, False, False)
    If Not foundCell Is Nothing Then
        v = foundCell.Value
        If IsNumeric(v) Then
            If CDbl(v) = 0 Then
                resultRow = foundCell.Row
                					  strStatus = "Pass"
							strMsg = "Row with Zero value found: """&excelPath&""""
							Call ReporterFunction(strLibraryFileName,"FindFirstRow_Zero","2","Row with Zero value found: ", strMsg)
            End If
        Else
            If Trim(CStr(v)) = "0" Then
                resultRow = foundCell.Row
                					 strStatus = "Pass"
							strMsg = "Row with Zero value found: """&excelPath&""""
							Call ReporterFunction(strLibraryFileName,"FindFirstRow_Zero","2","Row with Zero value found: ", strMsg)
            End If
        End If
    End If

    ' Fallback: loop through rows if Find didn't locate exact zero
    If resultRow = 0 Then
        Dim r, cellVal
        For r = startRow To lastRow
            cellVal = ws.Cells(r, colIndex).Value
            If IsNumeric(cellVal) Then
                If CDbl(cellVal) = 0 Then
                    resultRow = r
                   					  strStatus = "Pass"
							strMsg = "Row with Zero value found: """&excelPath&""""
							Call ReporterFunction(strLibraryFileName,"FindFirstRow_Zero","2","Row with Zero value found: ", strMsg)
                    Exit For
                End If
            Else
                If Trim(CStr(cellVal)) = "0" Then
                    resultRow = r
                     				strStatus = "Pass"
							strMsg = "Row with Zero value found: """&excelPath&""""
							Call ReporterFunction(strLibraryFileName,"FindFirstRow_Zero","2","Row with Zero value found: ", strMsg)
                    Exit For
                End If
            End If
        Next
    End If
    
    FindFirstRow_Zero_ByPath_NoCleanup = resultRow ' caller will close wb/xlApp
    
  	
       If blnDefault_eSwiftReporting Then  
		   Call UpdateResultHtml(strStepName,resultRow,strMsg,strStatus,"")
	   End If
  xlApp.Quit
 Set wb = Nothing
    Set xlApp = Nothing
    Set fso = Nothing
	   
End Function

Public Function SaveXmlAsXlsx(xmlPath, xlsxPath)
    Dim xlApp, wb, fso, bOK
    bOK = False
strStepName = "Conversion of XML file in to Excel file"
    On Error Resume Next

    ' Validate input files/paths
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(xmlPath) Then
        Reporter.ReportEvent micFail, "SaveXmlAsXlsx", "Source XML not found: " & xmlPath
        SaveXmlAsXlsx = False
        						strStatus = "FAIL"
							strMsg = "Conversion of XML to EXCEL: """&xmlPath&""" to the location: """&xlsxPath&""""
							Call ReporterFunction(strLibraryFileName,"Conversion of XML to EXCEL","1","Conversion of XML to EXCEL not copleted", strMsg)
        Exit Function
    End If

    ' Ensure target folder exists
    Dim targetFolder
    targetFolder = fso.GetParentFolderName(xlsxPath)
    If Len(targetFolder) > 0 And Not fso.FolderExists(targetFolder) Then
        fso.CreateFolder targetFolder
    End If

    ' Launch or bind to Excel
    Set xlApp = CreateObject("Excel.Application")
    If Err.Number <> 0 Or IsEmpty(xlApp) Or xlApp Is Nothing Then
        Reporter.ReportEvent micFail, "SaveXmlAsXlsx", "Could not create Excel.Application. Error: " & Err.Description
        strStatus = "FAIL"
							strMsg = "Conversion of XML to EXCEL: """&xmlPath&""" to the location: """&xlsxPath&""""
							Call ReporterFunction(strLibraryFileName,"Conversion of XML to EXCEL","1","Conversion of XML to EXCEL not copleted", strMsg)
        Err.Clear
        SaveXmlAsXlsx = False
        Exit Function
    End If

    ' Optional: keep Excel hidden for speed
    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    ' Open XML file - Excel can open SpreadsheetML 2003 .xml directly
    Set wb = xlApp.Workbooks.Open(xmlPath)
    If Err.Number <> 0 Or wb Is Nothing Then
        Reporter.ReportEvent micFail, "SaveXmlAsXlsx", "Failed to open XML in Excel. Error: " & Err.Description
        						strStatus = "FAIL"
							strMsg = "Conversion of XML to EXCEL: """&xmlPath&""" to the location: """&xlsxPath&""""
							Call ReporterFunction(strLibraryFileName,"Conversion of XML to EXCEL","1","Conversion of XML to EXCEL not copleted", strMsg)
        Err.Clear
        
    End If

    ' Define FileFormat constants for late binding
    ' xlOpenXMLWorkbook = 51 (.xlsx)
    Const xlOpenXMLWorkbook = 51
    ' Optional alternatives:
    ' xlWorkbookDefault = 51
    ' xlOpenXMLWorkbookMacroEnabled = 52 (.xlsm)

    ' If target exists, delete to avoid SaveAs prompt issues
    If fso.FileExists(xlsxPath) Then
        On Error Resume Next
        fso.DeleteFile xlsxPath, True
        On Error GoTo 0
    End If

    ' Save as .xlsx
    wb.SaveAs xlsxPath, xlOpenXMLWorkbook
   							 strStatus = "Pass"
							strMsg = "Conversion of XML to EXCEL: """&xmlPath&""" to the location: """&xlsxPath&""""
							Call ReporterFunction(strLibraryFileName,"Conversion of XML to EXCEL","2","Conversion of XML to EXCEL copleted", strMsg)
    If Err.Number <> 0 Then
        Reporter.ReportEvent micFail, "SaveXmlAsXlsx", "SaveAs failed. Error: " & Err.Description
        
        strStatus = "FAIL"
							strMsg = "Conversion of XML to EXCEL: """&xmlPath&""" to the location: """&xlsxPath&""""
							Call ReporterFunction(strLibraryFileName,"Conversion of XML to EXCEL","1","Conversion of XML to EXCEL not copleted", strMsg)
        Err.Clear
        
    End If

    bOK = True
    Reporter.ReportEvent micDone, "SaveXmlAsXlsx", "Saved successfully: " & xlsxPath
  If blnDefault_eSwiftReporting Then  
		   Call UpdateResultHtml(strStepName,xlsxPath,strMsg,strStatus,"")
	   End If
    ' Close workbook and Excel
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xlApp Is Nothing Then
        xlApp.DisplayAlerts = True
        xlApp.Quit
    End If
    Set wb = Nothing
    Set xlApp = Nothing
    Set fso = Nothing
    On Error GoTo 0

    SaveXmlAsXlsx = bOK
End Function

