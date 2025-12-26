'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.6.1.7.1. Check Price Change File from Workflow
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_2.6.1.7.1. Check Price Change File from Workflow"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.6.1.7.1. Check Price Change File from Workflow.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode VK11 ----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_VK11_0100_CONDITION_TYPE,False)
Call TakeScreenShot()

Call ClickButton("Key Combination   \(Shift\+F5\)",False)
Call TakeScreenShot()

Call ClickButton("Choose   \(Enter\)",True)
Call TakeScreenShot()

Call SetTextbox("Country","KOMG-LAND1","",DT_VK11_1990_COUNTRY,False)

Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Article",1,DT_VK11_1990_TABLECELL_ARTICLE_0,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Sales Unit",1,DT_VK11_1990_TABLECELL_SALES_UNIT_0,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Amount",1,DT_VK11_1990_TABLECELL_AMOUNT_0,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Unit",1,DT_VK11_1990_TABLECELL_UNIT_0,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","per",1," ",False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Valid From",1,ConvertDate(DT_VK11_1990_TABLECELL_VALID_FROM_0),False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","valid to",1,ConvertDate(DT_VK11_1990_TABLECELL_VALID_TO_0),False)

Call PressEnter()
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call VerifyStatusBar(DT_VK11_1990_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus ()


