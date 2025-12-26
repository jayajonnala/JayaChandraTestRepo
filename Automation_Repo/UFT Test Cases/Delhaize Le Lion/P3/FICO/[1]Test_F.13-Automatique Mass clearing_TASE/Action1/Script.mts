'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_F.13-Automatique Mass clearing  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_F.13-Automatiq Mass clear"
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
'
''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''--------TransactionCode-F.13----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE) 

Call SelectCheckbox("X_SAKNR", 0,DT_F13_1000_SELECT_GL_ACCOUNTS, False)
Call SetTextbox("Company Code","BUKRX-LOW","",DT_F13_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","GJAHX-LOW","",DT_F13_1000_FISCAL_YEAR,False)
Call ClickButtonIfExist("%_KONTS_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value", 1,"" , "",DT_F13_1000_GL_ACCOUNTS, True)
Call TakeScreenShot

Call ClickButtonIfExist("Copy   \(F8\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("%_DOCNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value", 1,"" , "",DT_F13_3010_TABLECELL_SINGLE_VALUE_0, True)
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

'Call ClickButtonIfExist("Last Page   \(Ctrl\+Page\ down\)",False)
'Call TakeScreenShot
Call VerifyifGuiLabelExists("No errors were logged during clearing in test run")
Call VerifyifGuiLabelExists_ByIndex("0,00",0)

Call ClickButtonIfExist("Back   \(F3\)",False)
Call SelectCheckbox("X_TESTL", 0,DT_F13_1000_TEST_RUN, False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call TakeScreenShot
Call PressEnter()
Call GetStatusBar("item1", "DT_F13_0120_CHECK_MESSAGEPARAMETER_OF_STATUSBAR")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",4)
'Call VerifyStatusBar(DT_F13_0120_CHECK_TEXT_OF_STATUSBAR)
Call ClickButtonIfExist("Back   \(F3\)",False)
Call ClickButtonIfExist("Back   \(F3\)",False)

'''''''--------TransactionCode-/faglb03----------''''

Call SetTcode(DT_F13_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Account Number","RACCT-LOW","",DT_F13_1000_ACCOUNT_NUMBER,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_F13_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_F13_1000_FISCAL_YEAR_OCC1,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call DoubleClickGuiGridCell("","",DT_MONTH,"Period",False)
Call TakeScreenShot
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call ClickButtonToolBar("&FIND",0)
'Call ClickButton("Find",True)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","DOCUMENT NUMBER",True)
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call ClickButton("Define Filter Values",True)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_F13_1105_DOCUMENT_NUMBER,True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
'Call SetTextbox("Find","GD_SEARCHSTR","","DOCUMENT NUMBER",True)
'Call ClickButton("Continue   \(Enter\)",True)
'Call ClickButton("Show sel. fields \(CTRL\+F3\)",True)
'Call ClickButton("Copy   \(Enter\)",True)
'Call SetTextbox("Document Number","%%DYN001-LOW","",DT_F13_1105_DOCUMENT_NUMBER,True)
Call TakeScreenShot
'Call ClickButton("Execute   \(Enter\)",True)
'Call TakeScreenShot
'Call ClickButtonIfExist("Change layout...   \(Ctrl\+F8\)",False)
'Call SelectRowGuiGrid("Column Set","","Column Name","Document Number",True)
'Call ClickButtonIfExist("APP_WL_SING",True)
'Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call VerifyGridCellContentbyName("shell",1,"Document Number","",DT_F13_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContentbyName("shell",1,"Amount in local currency","",DT_F13_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContentbyName("shell",2,"Amount in local currency","",DT_F13_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)
'Call VerifyifGuiLabelExists_ByIndex(DT_F13_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ICO_AUGP,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_F13_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_F13_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_F13_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB,0)
Call TakeScreenShot
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
