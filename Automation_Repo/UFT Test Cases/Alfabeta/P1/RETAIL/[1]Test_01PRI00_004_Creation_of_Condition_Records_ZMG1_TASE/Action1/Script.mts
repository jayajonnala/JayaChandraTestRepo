

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_01PRI00_004_Creation_of_Condition_Records_ZMG1
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


gstrTestCaseName = "Test_01PRI00_004_Creation_of_Condition_Records_ZMG1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_004_Creation_of_Condition_Records_ZMG1_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =4
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'--------------------------------------------  VK11------------------------------------------------
Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_VK11_100_CONDITION_TYPE,False)
Call TakeScreenShot()
Call PressEnter()

Call SelectRadioButton("RV130-SELKZ","Article per SOrg/DstCh",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Choose   \(Enter\)",True)
Call TakeScreenShot()

Call SetTextbox("Sales Organization","KOMG-VKORG","",DT_VK11_1073_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","KOMG-VTWEG","",DT_VK11_1073_DISTRIBUTION_CHANNEL,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Article",1,"","",DT_VK11_1073_TABLECELL_ARTICLE_0,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Sales unit",1,"","",DT_VK11_1073_TABLECELL_SALES_UNIT_0,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Amount",1,"","",DT_VK11_1073_TABLECELL_AMOUNT_0,False)
Call PressEnter()
Call TakeScreenShot()


Call ClickButton("Save   \(Ctrl\+S\)",False) 
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call VerifyStatusBarMessageType("S")
Call VerifyStatusBar(DT_VK11_1073_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus ()



