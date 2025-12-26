'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Clear GL Accounts Manual and Automatic_p7_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 22th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Clear GL Accounts Manual and Automatic_p7_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear GL Accounts  Manual and Automatic_p7_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''----------------------Tcode F.13----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SelectCheckbox("X_SAKNR","1",DT_F13_1000_SELECT_GL_ACCOUNTS,False)

Call SetTextbox("Company Code","BUKRX-LOW","",DT_F13_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","GJAHX-LOW","",DT_F13_1000_FISCAL_YEAR,False)
Call SetTextbox("Document Number","DOCNR-LOW","",DT_F13_1000_DOCUMENT_NUMBER,False)
Call SetTextbox("to","DOCNR-HIGH","",DT_F13_1000_TO,False)
Call SetTextbox("G/L Accounts","KONTS-LOW","",DT_F13_1000_GL_ACCOUNTS,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call SelectCheckbox("X_TESTL","1",DT_F13_1000_TEST_RUN,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

Call PressEnter()   
Wait(5)
'Capture the screenshot
Call TakeScreenShot()

'Validate If doc is generated
Call GetStatusBar("item1","DT_F13_0120_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_F13_0120_CHECK_TEXT_OF_STATUSBAR)

Call VerifyifGuiLabelExists(DT_F13_0120_CHECK_TEXT_OF_NO_NAME)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

