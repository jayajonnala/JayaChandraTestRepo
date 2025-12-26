
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_159_Update Account Completion Table (All Opcos)_TASE
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
'----------------------------------------------------------------------------------------------------------------------------
gstrTestCaseName = "TC1_Test_02-06-01-12-maintain fut"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Login 
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'-------------------------VL10G----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_MEK1_0100_CONDITION_TYPE,False)
Call TakeScreenShot()
Call PressEnter()
Call SelectRadioButton("RV130-SELKZ","Purch\.Org\./Vendor/Article",True)
Call TakeScreenShot()
Call ClickButton("Choose   \(Enter\)",True)
Call TakeScreenShot()
Call SetTextbox("Purch\. Organization","KOMG-EKORG","",DT_MEK1_1992_PURCH_ORGANIZATION,False)
Call SetTextboxNoLabel("KOMG-LIFNR","",DT_MEK1_1992_VENDOR,False)
Call TakeScreenShot()
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Article",1,"","",DT_MEK1_1992_TABLECELL_ARTICLE_0,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Amount",1,"","",DT_MEK1_1992_TABLECELL_AMOUNT_0,False)
Call TakeScreenShot()
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Valid from",1,"","",DT_MEK1_1992_TABLECELL_VALID_FROM_0,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Valid to",1,"","",DT_MEK1_1992_TABLECELL_VALID_TO_0,False)
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call TakeScreenShot()
Call VerifyStatusBar(DT_MEK1_1992_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus()

