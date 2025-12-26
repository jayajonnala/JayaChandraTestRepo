'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_02-06-02-04-17-ManualSalesPriceCalcExceptionRules_ArtPriceList_TASE
'.................Test Scenario: AT_04-06-02-06-04-Vendor returns-DC Drinks-No PO-Tied Empties
'.................TCode: ME21N,ME23N
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

gstrTestCaseName = "Test_02-06-02-04-17-ManualSaPriceList"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.

'//----------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------//

'Call CloseSessionsSAP()
'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)'.......................Mandatory Initial Call only in First Component in a Test Scenario

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'TCode ME21N
Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter()
Call TakeScreenShot()

Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_VK11_0100_CONDITION_TYPE,False)
Call PressEnter()
Call TakeScreenSHot()
Call ClickButton("Choose   \(Enter\)",True)
Call TakeScreenSHot()
Call SetTextbox("Sales Organization","KOMG-VKORG","",DT_VK11_1155_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","KOMG-VTWEG","",DT_VK11_1155_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Price List","KOMG-PLTYP","",DT_VK11_1155_PRICE_LIST,False)
Call TakeScreenSHot()
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Article",1,DT_VK11_1155_TABLECELL_ARTICLE_0,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Sales Unit",1,DT_VK11_1155_TABLECELL_SALES_UNIT_0,False)
Call PressEnter()
Call TakeScreenSHot()
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Amount",1,DT_VK11_1155_TABLECELL_AMOUNT_0,False)
Call PressEnter()
Call TakeScreenSHot()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenSHot()
Call VerifyStatusBar(DT_VK11_1155_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus()


