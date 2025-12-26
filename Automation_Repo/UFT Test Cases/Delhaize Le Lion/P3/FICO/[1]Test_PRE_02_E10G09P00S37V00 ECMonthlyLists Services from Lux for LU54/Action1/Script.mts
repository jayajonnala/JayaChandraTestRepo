'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_PRE_02_E10G09P00S37V00 ECMonthlyLists Services from Lux for LU54
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_PRE_02_E10G09P00S37V00  Lux for LU54"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

'Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''--------TransactionCode-ZFIGL_UPLOAD_POST----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call RefreshExcelSheet(DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME)
Call SetTextbox("File path name","P_FILE","",DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
wait 10
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButton("Yes",True)
Call TakeScreenShot

''''''--------TransactionCode--SM35---------''''

Call SetTcode(DT_ZFIGL_UPLOAD_POST_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call PressEnter()
Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1, False)
Call ClickButton("Process session   \(F8\)",False)
Call SelectRadioButton("D0300-BATCH", "Background", True)
Call TakeScreenShot
Call ClickButton("Process   \(Enter\)",True)
Call TakeScreenShot
wait 30
Call ClickButton("Analyze session   \(F2\)",False)
wait 70
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
wait 50
Call ClickButton("Analyze session   \(F2\)",False)

' VerifyTextBoxContent(textboxAttachedText, textboxName, textboxIndex, expectedValue, blnIsItPopup)
Call VerifyTextBoxContent("Processed","APQDCNT-TRANSCNTF",0,DT_INDEX_2,False)
Call VerifyTextBoxContent("With Errors","APQDCNT-TRANSCNTE",0,DT_INDEX_1,False)
'Call VerifyTextBoxContent("With Errors","APQDCNT-TRANSCNTE",0,DT_INDEX_2,False)

'Call FindRowNumber("RSBDC_ANALYSETC_TCODES", "", "",DT_GET_ROW)
Call SelectTab("TAB_DYNPRO"," Log created on "&ConvertDate(DT_SM35_1005_TO__OCC1),False)
Call SelectRadioButton("RB-PRO_TCODE", "Transaction", False)

Call FindRowNumber("RSBDC_ANALYSETC_PROTOCOL", "No.", "312","DT_GET_DOC_ROW_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_GET_DOC_ROW_OUTPUT",DT_GET_DOC_ROW)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)

Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message",DT_GET_DOC_ROW, "", "","DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_3_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_3_OUTPUT",DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_3)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call TakeScreenShot

' FindRowNumber(tableNameOrGridTitle, refColumnName, refCellValue, dataTableColumnName)
' GetTableCellData(tableName, columnName, rowNumber, refColumnName, refCellValue, dataTableColumnName, blnIsItPopup)

Call ClickButton("Last transaction   \(Ctrl\+F6\)",False)
Call TakeScreenShot

Call FindRowNumber("RSBDC_ANALYSETC_PROTOCOL","No.", "312","DT_GET_DOC_ROW_OCC1_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_GET_DOC_ROW_OCC1_OUTPUT",DT_GET_DOC_ROW_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)

Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message",DT_GET_DOC_ROW_OCC1, "", "","DT_OP_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_7_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_OP_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_7_OUTPUT",DT_OP_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_7)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)

Call TakeScreenShot
Call SelectRadioButton("RB-PRO_SESSION", "Folder", False)
Call TakeScreenShot
Call VerifyTableCellContent(5, "Message", "RSBDC_ANALYSETC_PROTOCOL", DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_12)
Call VerifyTableCellContent(4, "Message", "RSBDC_ANALYSETC_PROTOCOL", DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_11)

'GetRowNo =4
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
'Call TakeScreenShot

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
