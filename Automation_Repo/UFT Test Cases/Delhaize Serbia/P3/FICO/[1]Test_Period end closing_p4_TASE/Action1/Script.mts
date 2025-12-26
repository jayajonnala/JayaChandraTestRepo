
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Period end closing_p4
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
'.................Test Script Name :Test_Period end closing_p4
'.................Author : TCS 	   :Raushan
'................ Creation Date    :16th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Period end closing_p4"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.7.1.2.1. Create Low Level Header ZRPC.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode KSV5----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call SetTextbox("Period","RKGA2U-FROM","",DT_KSV5_0101_PERIOD,FALSE)
Call SetTextbox("To","RKGA2U-TO","",DT_KSV5_0101_TO,FALSE)
Call SetTextbox("Fiscal Year","RKGA2U-GJAHR","",DT_KSV5_0101_FISCAL_YEAR,FALSE)


Call SetTextbox("Cycle","RKGA2-KSCYC","",DT_KSV5_0101_CYCLE,FALSE)
Call TakeScreenShot()
Call PressEnter()  

'Click on Execute
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

'Click on Line Item
Call ClickButton("Line items   \(F9\)",False) 
Call TakeScreenShot()

'Verify Grid Content
Call VerifyGridCellContent("",1,"Object",0,DT_KSV5_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_OBJID)
Call VerifyGridCellContent("",2,"Object",0,DT_KSV5_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_OBJID)

'Click on Back
Call ClickButtonIfExist("Back   \(F3\)",False)
wait(1)

Call ClickButton("Yes",True)
Wait(1)

'Click on Journal
Call ClickButton("Journal   \(Shift\+F6\)",False)
Call TakeScreenShot()

'Click on Back
Call ClickButtonIfExist("Back   \(F3\)",False)
wait(1)

Call ClickButton("Yes",True)
Wait(1)

'Click on Back
Call ClickButtonIfExist("Back   \(F3\)",False)
wait(1)

Call ClickButton("Yes",True)
Wait(1)

Call SelectCheckbox("RKGA2U-TEST",0,DT_KSV5_0101_TEST_RUN,False)
Call TakeScreenShot()

'Click on Execute
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

Call VerifyifGuiLabelExists(DT_KSV5_0120_CHECK_TEXT_OF_PROCESSING_COMPLETED_WITH_WARNINGS)



'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

