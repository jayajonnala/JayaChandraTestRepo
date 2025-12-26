'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_02-06-01-11-Maintain exception rule for Markup_TASE
'.................Test Scenario: AT_Wholesale price with exception rules
'.................TCode: VK11
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
	DataRowSet= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


gstrTestCaseName = "TC2_Test_02-06-01-11-Maint Markup"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.

'//----------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------//

'Call CloseSessionsSAP()
'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)'.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'---------------------VK11---------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter()
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextBox("Condition Type","RV13A-KSCHL","",DT_VK11_0100_CONDITION_TYPE,False)
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Choose   \(Enter\)",True)

Call SetTextBox("Sales Organization","KOMG-VKORG","",DT_VK11_1950_SALES_ORGANIZATION,False)
Call SetTextBox("Distribution Channel","KOMG-VTWEG","",DT_VK11_1950_DISTRIBUTION_CHANNEL,False)
Call TakeScreenShot()
Call SetTextBox("Price List","KOMG-PLTYP","",DT_VK11_1950_PRICE_LIST,False)
Call PressEnter()
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Article","1","","",DT_VK11_1950_TABLECELL_ARTICLE_0,False)
Call PressEnter()
Call TakeScreenShot()
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Amount","1","","",DT_VK11_1950_TABLECELL_AMOUNT_0,False)
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call VerifyStatusBar(DT_VK11_1950_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

Call LogOff()
Call FinalStatus()
