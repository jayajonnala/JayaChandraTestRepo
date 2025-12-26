'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_ZFIGL_UPLOAD_POST-SM35-Accruel with WBS and Order and Trading Pa  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_ZFIGL_UPLOAD_POST-SM35-Accruel"
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)



'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

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
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButton("Yes",True)
Call TakeScreenShot

''''''--------TransactionCode--SM35---------''''

Call SetTcode(DT_ZFIGL_UPLOAD_POST_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)


Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", 1, False)
Call ClickButton("Process session   \(F8\)",False)
Call TakeScreenShot
Call SelectRadioButton("D0300-ERROR", "Display errors only", True)
Call TakeScreenShot
Call ClickButton("Process   \(Enter\)",True)
Call TakeScreenShot
Wait 60
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Call TakeScreenShot
Call VerifyTextBoxContent("Information Message", "MESSTXT1", "", lcase(DT_ZFIGL_UPLOAD_POST_0010_CHECK_TEXT_OF_MESSTXT1), True)
Call ClickButton("Go back to batch input session overview   \(Enter\)",True)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
Call TakeScreenShot
Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", 1, False)
Call TakeScreenShot
Wait 60
Call ClickButton("Analyze session   \(F2\)",False)
Call TakeScreenShot
Call SelectTab("TAB_DYNPRO"," Log created on "&ConvertDate(DT_SM35_1000_LOG_DATE_OCC1),False) 'Need to check the tab
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message",3, "", "", "DT_OP_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_1", False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message", 5, "", "", "DT_OP_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_2", False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message", 7, "", "", "DT_OP_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_3", False)
Call VerifyTableCellContent(11, "Message", "RSBDC_ANALYSETC_PROTOCOL", DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_7)
Call TakeScreenShot
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
