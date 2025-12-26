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

'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Compensations Process_p6_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 7th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Compensations Process_p6_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Compensations Process_p6_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'''Login'''
'DataRowSet=2
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  


'----------------------Tcode F.64----------------------------


Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()  
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call SetTextbox("Correspondence","EVENT-LOW","",DT_F64_1000_CORRESPONDENCE,False)
Call SetTextbox("Company code","BUKRS-LOW","",DT_F64_1000_COMPANY_CODE,False)
Call SetTextbox("Document number","BELNR-LOW","",DT_F64_1000_DOCUMENT_NUMBER,False)
Call TakeScreenShot()

'let the bellow line disabled unless there are multiple records, if script is failing due to multiple record change the time in datasheet
'Call SetTextbox("Time of request","UZEIT-LOW","",DT_F64_1000_TIME,False)

Call ClickButton("Execute   \(F8\)",False)
Wait(1)
Call TakeScreenShot()

Call VerifyifGuiLabelExists(DT_F64_0120_CHECK_TEXT_OF_NO_NAME)

Call ClickLabel(DT_F64_0120_CHECK_TEXT_OF_NO_NAME,"",False)
Wait(1)
Call TakeScreenShot()

Call SetTextbox("Output Device","USR01-SPLD","",DT_F64_1100_OUTPUT_DEVICE,True)
Call TakeScreenShot()

Call ClickButton("Continue   \(Enter\)",True)
Wait(3)
Call TakeScreenShot()

Call SetTextbox("Output Device","USR01-SPLD","",DT_F64_1100_OUTPUT_DEVICE_OCC1,True)
Call TakeScreenShot()

Call ClickButton("Continue   \(Enter\)",True)
Wait(4)
Call TakeScreenShot()

Wait(2)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()


