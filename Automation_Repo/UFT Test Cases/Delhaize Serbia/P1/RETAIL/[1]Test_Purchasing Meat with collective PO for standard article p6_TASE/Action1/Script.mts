
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Purchasing Meat with collective PO for standard article p6
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
'.................Test Script Name :Test_Purchasing Meat with collective PO for standard article p6
'.................Author : TCS 	   :Raushan
'................ Creation Date    :16th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Purchasing Meat with collective PO for standard article p6"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Distribution process to DS stores - SW41 (Fresh meat) - p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode MIGO----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_DOCUMENT_NO,False)
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",DT_MIR4_6150_FISCAL_YEAR,False)
Call TakeScreenShot
Call ClickButton("Display Document   \(F2\)",false)
Call TakeScreenShot
Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",false)
Call TakeScreenShot

Call VerifyWindowValue(DT_MIR4_0750_CHECK_TEXT_OF_TITL)
Call VerifyTextBoxContent("Fiscal Year","BKPF-GJAHR","",DT_MIR4_0750_CHECK_TEXT_OF_FISCAL_YEAR,False)
Call VerifyTextBoxContent("Company Code","BKPF-BUKRS","",DT_MIR4_0750_CHECK_TEXT_OF_COMPANY_CODE,False)

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot

Call LogOff()
Call FinalStatus ()








'*********************************************End Of Script*********************************************************************

