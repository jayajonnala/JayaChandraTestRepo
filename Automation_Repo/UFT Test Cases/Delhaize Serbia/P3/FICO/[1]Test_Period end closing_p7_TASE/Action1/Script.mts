
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Period end closing_p7
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
'.................Test Script Name :Test_Period end closing_p7
'.................Author : TCS 	   :Raushan
'................ Creation Date    :16th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Period end closing_p7"
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

'Enter the details
Call SetTextbox("Period","RKGA2U-FROM","",DT_KSV5_0101_PERIOD,FALSE)
Call SetTextbox("To","RKGA2U-TO","",DT_KSV5_0101_TO,FALSE)
Call SetTextbox("Fiscal Year","RKGA2U-GJAHR","",DT_KSV5_0101_FISCAL_YEAR,FALSE)
Call SetTextbox("Cycle","RKGA2-KSCYC","",DT_KSV5_0101_CYCLE,FALSE)
Call SelectCheckbox("RKGA2U-TEST",0,"OFF",False)
Call TakeScreenShot()

'Select Distribution/Reverse on Menu Bar
Call SelectMenuBar("Distribution;Reverse")

'Verify the Text box message
Call VerifyTextBoxContent("Information Message","MESSTXT1",0,Lcase(DT_KSV5_0010_CHECK_TEXT_OF_MESSTXT1),True)

Call ClickButton("Continue   \(Enter\)",True) 

'----------------------Tcode KSV6----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_KSV5_0101_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_KSV5_0101_OKCD)

Call PressEnter() 

'Click on Execute
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

Call VerifyGridCellContent("",1,"Posting Period",0,DT_KSV6_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PERIOD)
Call VerifyGridCellContent("",1,"Cycle",0,DT_KSV6_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_CYCLE)

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

