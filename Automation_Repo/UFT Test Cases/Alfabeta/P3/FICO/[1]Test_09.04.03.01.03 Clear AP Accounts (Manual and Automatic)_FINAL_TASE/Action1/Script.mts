

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.03.01.03 Clear AP Accounts (Manual and Automatic)_FINAL
'.................Author : TCS 
'................ Creation Date :
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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrTestCaseName = "Test_09.04.03.01.03 Clear AP Accounts (Manual and Automatic)_FINAL"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\jjonn\Desktop\TASEWork\Data\TASE_DT_09.04.01.01.01 Manage Manual Post  Direct Domestic Vendor Invoic.xls"
'strResultFolderPath = "C:\Users\jjonn\Desktop\TASEWork\Results"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''--------TransactionCode-F.13 ----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

DT_F13_1000_Assignment = year(DT_F13_1000_FISCAL_YEAR)& ConvertDoubledigit(CSTR(Month(DT_F13_1000_FISCAL_YEAR)))& ConvertDoubledigit(CSTR(Day(DT_F13_1000_FISCAL_YEAR)))

Call SetTextbox("Company Code","BUKRX-LOW","",DT_F13_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","GJAHX-LOW","",Year(DT_F13_1000_FISCAL_YEAR),False)
'Call SetTextbox("Assignment","SO_ZUONR-LOW","",DT_F13_1000_Assignment,False)


Call ClickButton("%_DOCNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F13_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_F13_3010_TABLECELL_SINGLE_VALUE_1,True)
Call ClickButton("Copy   \(F8\)",True)

'Call SetTextboxNoLabel("KONTK-LOW","",DT_FB60_0010_VENDOR,False)
'Call SetTextbox("Vendors","KONTK-LOW","",DT_FB60_0010_VENDOR,False)
Call SetTextbox("Posting Date","POSTDATE-LOW","",DT_F13_1000_POSTING_DATE,False)

Call SelectCheckbox("X_LIFNR", 0, "ON", False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

'Call VerifyifGuiLabelExists(DT_F13_0120_CHECK_TEXT_OF_NO_NAME)
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call SelectCheckbox("X_TESTL", 0, "OFF", False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot

Call GetStatusBar("item1","DT_F13_0120_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_F13_0120_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_F13_0120_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_F13_0120_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call TakeScreenShot
Call ClickButton("Exit   \(Shift\+F3\)",False)
Call ClickButton("Exit   \(Shift\+F3\)",False)

'''''--------TransactionCode-FBL1N ----------''''
Call SetTcode(DT_F13_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectRadioButton("X_CLSEL","Cleared items",False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN012_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F13_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_F13_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","3","","",DT_F13_3010_TABLECELL_SINGLE_VALUE_2,True)

Call ClickButton("Copy   \(F8\)",True)
Call SetTextbox("Clearing date", "SO_AUGDT-LOW", 0, ConvertDate(DT_TODAY), False)

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyStatusBarMessageType("S")

Call VerifyGridCellContent("", 1, "ICO_AUGP", 0, DT_F13_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ICO_AUGP)
Call VerifyGridCellContent("", 5, "DMSHB", 0, DT_F13_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_DMSHB)

Call LogOff'
Call FinalStatus()




