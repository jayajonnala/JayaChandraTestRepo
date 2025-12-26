'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Clear AP Accounts - Manual_p7_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 9th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Clear AP Accounts - Manual_p7_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear AP Accounts - Manual_p7_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode F.13----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Company Code","BUKRX-LOW","",DT_F13_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","GJAHX-LOW","",DT_F13_1000_FISCAL_YEAR,False)

'Capture the screenshot
Call TakeScreenShot()

SapGuiSession("transaction:=F\.13").SapGuiWindow("transaction:=F\.13").SAPGuiButton("tooltip:=Multiple selection","index:=1").Click
'Capture the screenshot
Call TakeScreenShot()

'set filter criteria
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_F13_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_F13_3010_TABLECELL_SINGLE_VALUE_1,True)

'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Copy   \(F8\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call SelectCheckbox("X_LIFNR","1",DT_F13_1000_SELECT_VENDORS,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Clearing date","AUGDT","",Replace(DT_F13_1000_CLEARING_DATE,"/","."),False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectCheckbox("X_TESTL","1",DT_F13_1000_TEST_RUN,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(1)
Call TakeScreenShot()

Call PressEnter()   
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

'Validate If doc is generated
Call GetStatusBar("item1","DT_F13_0120_DOCUMENT_NUMBER_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_F13_0120_CHECK_TEXT_OF_STATUSBAR)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

