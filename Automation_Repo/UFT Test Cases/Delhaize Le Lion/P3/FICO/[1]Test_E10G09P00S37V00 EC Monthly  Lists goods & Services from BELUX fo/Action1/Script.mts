'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_E10G09P00S37V00 EC Monthly  Lists goods & Services from BELUX fo  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_E10G09P00S37V00  Services from BELUX fo"
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
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''--------TransactionCode-/ns_alr_87012402----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SelectCheckbox("PAR_SRV",0,"ON", False)
Call SelectCheckbox("PAR_EPOS",0,"ON", False)
Call TakeScreenShot
Call SetTextbox("Company Code","BR_BUKRS-LOW","",DT_S_ALR_87012402_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","",DT_S_ALR_87012402_1000_FISCAL_YEAR,False)
Call SetTextbox("Posting date","BR_BUDAT-LOW","",ConvertDate(DT_S_ALR_87012402_1000_POSTING_DATE),False)
Call SetTextbox("to","BR_BUDAT-HIGH","",ConvertDate(DT_S_ALR_87012402_1000_TO),False)
Call SetTextbox("Reporting country","SEL_LSTM-LOW","",DT_S_ALR_87012402_1000_REPORTING_COUNTRY,False)
Call SetTextbox("Output Tax Code","SEL_MWKZ-LOW","","",False)
Call ClickButton("%_BR_BELNR_%_APP_%-VALU_PUSH", False)
Call SelectTab("TAB_STRIP", "Select Ranges", True)
Call TakeScreenShot

Call SetTableData("SAPLALDBINTERVAL", "Upper limit", 1, "", "", DT_S_ALR_87012402_3020_TABLECELL_UPPER_LIMIT_0,True)
Call SetTableData("SAPLALDBINTERVAL", "Lower limit", 1, "", "", DT_S_ALR_87012402_3020_TABLECELL_LOWER_LIMIT_0,True)
'Call SetTableData("SAPLALDBINTERVAL", "Upper limit", 1, "", "", DT_S_ALR_87012402_3020_TABLECELL_UPPER_LIMIT_0,True)

Call ClickButton("Copy   \(F8\)", True)
Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot
Call ClickButton("Change layout...   \(Ctrl\+F8\)", False)
Call ClickButton("Find", True)
Call TakeScreenShot
Call SetTextbox("Find","GD_SEARCHSTR","","Tax Code",True)
Call ClickButton("Continue   \(Enter\)", True)
Call ClickButton("Show Sel. Fields \(Ctrl\+F3\)", True)

Call TakeScreenShot
Call ClickButton("Copy   \(Enter\)", True)
Call VerifyStatusBarMessageType("S")

Call ClickButton("Sort in Ascending Order   \(Ctrl\+F4\)", False)
Call ClickButton("Copy   \(Enter\)", True)
Call VerifyStatusBarMessageType("S")

Call ClickButton("Subtotals...   \(Ctrl\+Shift\+F6\)", False)
Call ClickButton("Find", True)
Call TakeScreenShot
Call SetTextbox("Find","GD_SEARCHSTR","","Tax Code",True)
Call ClickButton("Continue   \(Enter\)", True)
Call ClickButton("Show Sel. Fields \(Ctrl\+F3\)", True)
Call TakeScreenShot
Call ClickButton("Copy   \(Enter\)", True)
Call VerifyStatusBarMessageType("S")

Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012402_0120_CHECK_TEXT_OF_NO_NAME, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012402_0120_CHECK_TEXT_OF_NO_NAME_OCC1, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012402_0120_CHECK_TEXT_OF_NO_NAME_OCC2, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012402_0120_CHECK_TEXT_OF_NO_NAME_OCC3, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012402_0120_CHECK_TEXT_OF_NO_NAME_OCC4, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012402_0120_CHECK_TEXT_OF_NO_NAME_OCC5, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012402_0120_CHECK_TEXT_OF_NO_NAME_OCC6, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012402_0120_CHECK_TEXT_OF_NO_NAME_OCC7, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012402_0120_CHECK_TEXT_OF_NO_NAME_OCC8, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012402_0120_CHECK_TEXT_OF_NO_NAME_OCC9, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012402_0120_CHECK_TEXT_OF_NO_NAME_OCC10, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012402_0120_CHECK_TEXT_OF_NO_NAME_OCC11, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012402_0120_CHECK_TEXT_OF_NO_NAME_OCC12, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012402_0120_CHECK_TEXT_OF_NO_NAME_OCC13, 1)
Call ClickButton("Next page   \(Page down\)", False)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012402_0120_CHECK_TEXT_OF_NO_NAME_OCC14, 1)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012402_0120_CHECK_TEXT_OF_3A, 1)
Call ClickButton("Last Page   \(Ctrl\+Page down\)", False)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012402_0120_CHECK_TEXT_OF_NO_NAME_OCC15, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012402_0120_CHECK_TEXT_OF_NO_NAME_OCC16, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012402_0120_CHECK_TEXT_OF_NO_NAME_OCC17, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012402_0120_CHECK_TEXT_OF_NO_NAME_OCC18, 1)
Call TakeScreenShot

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
