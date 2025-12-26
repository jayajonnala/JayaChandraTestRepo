'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_ACACACT-Calculate and post accruals_TASE
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_ACACACT-Calculate post accruals"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''Call EndDateof445PeriodByDate((DT_DATE), "DT_ENDDATE_445_OUTPUT")
'''Call GetInputFromExcel(gstrInputExcelFilePathAndName, "Global", datatable_row)
'
''''''--------------login----------------'''''

'Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''''--------TransactionCode-/ACACACT----------''''

Call EndDateof445PeriodByDate((DT_DATE), "DT_ENDDATE_445_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName, "Global", datatable_row)

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)


Call SetTextbox("Accounting Principle","SO_ACRLE-LOW","",DT_ACACACT_1000_ACCOUNTING_PRINCIPLE,False)
Call SetTextbox("Accrual Object Number","SO_OBJNM-LOW","",DT_ACACACT_1000_ACCRUAL_OBJECT_NUMBER,False)
Call SetTextbox("Key Date for Accruals","P_EFDATE","",ConvertDate(DT_ACACACT_1000_KEY_DATE_FOR_ACCRUALS),False)
Call SetTextbox("Company Code","SO_BUKRS-LOW","",DT_ACACACT_1000_COMPANY_CODE,False)
Call SetTextbox("Execution Type","P_RUNTYP","",DT_ACACACT_1000_EXECUTION_TYPE,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButton("Display Messages Posted   \(F7\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Current display variant   \(Ctrl\+F8\)",True)
Call SetTableData("SAPLSKBHTC_WRITE_LIST","Length",3, "", "",45, True)
Call ClickButtonIfExist("Copy   \(Enter\)",True)
Call VerifyifGuiLabelExists("Document check - no errors: ACE.*")
Call VerifyifGuiLabelExists(".*R1ECLNT200")
Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Back   \(F3\)",False)
Call SelectCheckbox("P_TEST", 0, "OFF", False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call ClickButton("Display Messages Posted   \(F7\)",False)
Call TakeScreenShot
Call GetLabelContentByRefLabel("Message text",0,-48,"DT_OP_ACACACT_0120_CHECK_TEXT_OF_DOCUMENT_POSTED_SUCCESSFULLY_ACE_2000001901BE102015_R1ECLNT200", True)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",4)
'Call VerifyifGuiLabelExists(DT_ACACACT_0120_CHECK_TEXT_OF_DOCUMENT_POSTED_SUCCESSFULLY_ACE_2000001901BE102015_R1ECLNT200)
Call VerifyifGuiLabelExists("Processing.*")
Call VerifyifGuiLabelExists(".*carried out in update run")
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
