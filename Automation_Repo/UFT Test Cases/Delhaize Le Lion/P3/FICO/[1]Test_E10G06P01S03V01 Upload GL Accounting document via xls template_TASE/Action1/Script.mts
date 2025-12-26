'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_E10G06P01S03V01 Upload GL Accounting document via xls template  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_E10G06P01S03V01 Upload GL Acc"
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
	DataRowSet= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

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
Call SetTextbox("Session", "P_SESS", "",DT_ZFIGL_UPLOAD_POST_1000_SESSION,False)
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
Call SelectRadioButton("D0300-BATCH", "Background", True)
Call TakeScreenShot
Call ClickButton("Process   \(Enter\)",True)
Call TakeScreenShot
wait (60)
Call ClickButton("Analyze session   \(F2\)",False)
Call TakeScreenShot
wait (60)
Call ClickButton("Back   \(F3\)",False)
wait (60)
Call ClickButton("Analyze session   \(F2\)",False)
Call TakeScreenShot

Call VerifyTextBoxContent("Current Content", "APQDCNT-TRANSCNT", "", "12", False)
Call VerifyTextBoxContent("To Process", "APQDCNT-TRANSCNTO", "", "0", False)
Call SelectTab("TAB_DYNPRO"," Log created on "&ConvertDate(DT_SM35_1000_LOG_DATE_OCC1),False)


Call SelectRadioButton("RB-PRO_TCODE", "Transaction", False)
'''''transaction 1
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message",1, "", "","DT_OP_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_1", False)

Call TakeScreenShot
'''''transaction 2
Call ClickButton("Next Transaction   \(Ctrl\+F5\)",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message",1, "", "", "DT_OP_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_3", False)

Call TakeScreenShot
Call ClickButton("Next Transaction   \(Ctrl\+F5\)",False)
'''''transaction 3
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message",1, "", "","DT_OP_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_5", False)

Call TakeScreenShot
''''' transaction 4
Call ClickButton("Next Transaction   \(Ctrl\+F5\)",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message",4, "", "", "DT_OP_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_8", False)

Call TakeScreenShot
''''''transaction 6
Call ClickButton("Next Transaction   \(Ctrl\+F5\)",False)
Call ClickButton("Next Transaction   \(Ctrl\+F5\)",False)
''''Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message",1, "", "", "DT_OP_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_8", False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message",2, "", "", "DT_OP_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_16", False)

Call TakeScreenShot
''''''transaction 7
Call ClickButton("Next Transaction   \(Ctrl\+F5\)",False)
Call TakeScreenShot
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message",1, "", "", "DT_OP_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_18", False)

Call TakeScreenShot
''''''transaction 11
Call ClickButton("Last transaction   \(Ctrl\+F6\)",False)
Call ClickButton("Previous transaction   \(Ctrl\+F4\)",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message",1, "", "", "DT_OP_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_9", False)

Call TakeScreenShot
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
