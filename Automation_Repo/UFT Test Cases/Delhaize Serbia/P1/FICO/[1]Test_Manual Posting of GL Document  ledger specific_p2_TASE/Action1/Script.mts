
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Manual Posting of GL Document ledger specific_p2
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



'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manual Posting of GL Document ledger specific_p2
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 19th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manual Posting of GL Document ledger specific_p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Manual Posting of GL Document  ledger specific_p2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'Login to SAP System

'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode FB03----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 

'Enter the conpany code
Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB03_0100_DOCUMENT_NUMBER,False)   
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB03_0100_COMPANY_CODE,False)   
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB03_0100_FISCAL_YEAR,False)   
Wait(1)
Call TakeScreenShot()

'Click on Document List
Call ClickButton("Create Document List/Find Documents   \(Shift\+F8\)",False) 
Wait(1)
Call TakeScreenShot()

'Enter the details
Call SetTextbox("Company code","BR_BUKRS-LOW","",DT_FB03_1000_COMPANY_CODE,False)   
Call SetTextbox("Ledger","BR_RLDNR-LOW","",DT_FB03_1000_LEDGER,False)   
Call SetTextbox("Entry date","BR_CPUDT-LOW","",DT_FB03_1000_ENTRY_DATE,False)   
Call SelectCheckbox("UNAME",0,DT_FB03_1000_OWN_DOCUMENTS_ONLY,False)

Call PressEnter()
Wait(1)
Call TakeScreenShot()

'Click on Execute
Call ClickButton("Execute   \(F8\)",False) 
Wait(1)
Call TakeScreenShot()

Call VerifyStatusBarMessageType("S")


'Filter the Document No
Call SelectColumnGuiGrid("",0,"Document Number",False)
Call ClickButton("Set Filter   \(Ctrl\+F5\)",False) 
Wait(2)
Call TakeScreenShot()
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FB03_1105_DOCUMENT_NUMBER,True)  
Call TakeScreenShot()
Call ClickButton("Execute   \(Enter\)",True) 
Wait(1)
Call SelectRowGuiGrid("",0,"Document Number",DT_FB03_1105_DOCUMENT_NUMBER,False)
Call TakeScreenShot()

'Click on Choose Button
Call ClickButton("Choose   \(F2\)",False) 
Wait(1)
Call TakeScreenShot()

'Verify the Data
Call VerifyGridCellContent("",1,"Account",0,DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"Account",0,DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("",1,"Company Code",0,DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUKRS)

'Click on Back
Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(2)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

