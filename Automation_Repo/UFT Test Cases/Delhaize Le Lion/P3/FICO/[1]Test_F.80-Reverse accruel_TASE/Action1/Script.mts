'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_F.80-Reverse accruel 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_F.80-Reverse accruel"
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

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''--------TransactionCode-F.80----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


Call SetTextbox("Company code","BR_BUKRS-LOW","",DT_F80_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","",DT_F80_1000_FISCAL_YEAR,False)
Call SetTextbox("Ledger","BR_RLDNR-LOW","",DT_F80_1000_LEDGER,False)
Call TakeScreenShot
Call SetTextbox("Reason for reversal","STOGRD","",DT_F80_1000_REASON_FOR_REVERSAL,False)
Call ClickButton("\%\_BR\_BELNR\_\%\_APP\_\%\-VALU\_PUSH",False)
Call TakeScreenShot
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_F80_3010_TABLECELL_SINGLE_VALUE_0, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_F80_3010_TABLECELL_SINGLE_VALUE_1, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 3, "", "", DT_F80_3010_TABLECELL_SINGLE_VALUE_2, True)
Call ClickButton("Copy   \(F8\)",True)
Call TakeScreenShot
Call SetTextbox("Posting period","MONAT","",DT_F80_1000_POSTING_PERIOD,False) 
Call SetTextbox("Posting Date","STODAT","",ConvertDate(DT_F80_1000_POSTING_DATE),False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyifGuiLabelExists(DT_F80_0120_CHECK_TEXT_OF_NO_NAME)
Call VerifyifGuiLabelExists(DT_F80_0120_CHECK_TEXT_OF_NO_NAME_OCC1)
Call VerifyifGuiLabelExists(DT_F80_0120_CHECK_TEXT_OF_NO_NAME_OCC2)

Call ClickButton("Reverse documents   \(F6\)",False)


Call GetLabelContentByRefLabel("Message",0,-48,"DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_OUTPUT",DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_700009817)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)


Call VerifyifGuiLabelExists(DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_OCC1)
Call VerifyifGuiLabelExists(DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_OCC2)
Call VerifyifGuiLabelExists(DT_F80_0120_CHECK_TEXT_OF_NUMBER_OF_DOCUMENTS_REVERSED_3)

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
