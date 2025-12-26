
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.6.1.8.1. Maintain Wholesale Price (ZK07, ZK05)
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
'.................Test Script Name : Test_2.6.1.8.1. Maintain Wholesale Price (ZK07, ZK05)
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 24th Oct
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_2.6.1.8.1. Maintain Wholesale Price (ZK07, ZK05)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.6.1.8.1. Maintain Wholesale Price (ZK07, ZK05).xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
'DataRowSet=5
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode VK11 ----------------------------
'Enter the transaction code
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'Enter the Condition Type
Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_VK11_0100_CONDITION_TYPE,False)
'Capture the screenshot
Call TakeScreenShot()

'Click on Key Combination Button
Call ClickButton("Key Combination   \(Shift\+F5\)",False)
Wait(2)
Call TakeScreenShot()


'Click on Choose button in Pop Up
Call ClickButton("Choose   \(Enter\)",True)
Wait(2)
Call TakeScreenShot()

'Enter the details
Call SetTextbox("Sales Organization","KOMG-VKORG","",DT_VK11_1005_SALES_ORGANIZATION,False) 
Call SetTextbox("Distribution Channel","KOMG-VTWEG","",DT_VK11_1005_DISTRIBUTION_CHANNEL,False) 
Call SetTextbox("Customer","KOMG-KUNNR","",DT_VK11_1005_CUSTOMER,False) 


Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Article",1,DT_VK11_1005_TABLECELL_ARTICLE_0,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Amount",1,DT_VK11_1005_TABLECELL_AMOUNT_0,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","per",1," ",False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Valid From",1,ConvertDate(DT_VK11_1005_TABLECELL_VALID_FROM_0),False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","valid to",1,ConvertDate(DT_VK11_1005_TABLECELL_VALID_TO_0),False)

Call PressEnter()
wait(2)
'Capture the screenshot
Call TakeScreenShot()

'Save the Details
Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(2)
Call TakeScreenShot()

'Verify the status bar message
Call VerifyStatusBar(DT_VK11_1005_CHECK_TEXT_OF_STATUSBAR)


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

