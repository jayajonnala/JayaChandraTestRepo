'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_E10G09P00S37V00 EC Monthly  Lists Services from Lux for LU54  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_E10G09P00S37V00  Lux for LU54"
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

'Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''--------TransactionCode-S_ALR_87012357----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Company Code","BR_BUKRS-LOW","",DT_S_ALR_87012357_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","",DT_S_ALR_87012357_1000_FISCAL_YEAR,False)
Call SetTextbox("Posting date","BR_BUDAT-LOW","",ConvertDate(DT_S_ALR_87012357_1000_POSTING_DATE),False)
Call SetTextbox("to","BR_BUDAT-HIGH","",ConvertDate(DT_S_ALR_87012357_1000_TO),False)
Call ClickButton("PUSHB_O1", False)
'Call SetTextbox("Tax return country","SEL_LSTM-LOW","",DT_S_ALR_87012357_1000_TAX_RETURN_COUNTRY,False)
Call SetTextboxNoLabel("SEL_LSTM-LOW","",DT_S_ALR_87012357_1000_TAX_RETURN_COUNTRY,False)
Call ClickButton("%_BR_BELNR_%_APP_%-VALU_PUSH", False)
Call PressEnter()

Call SelectTab("TAB_STRIP", "Select Ranges", True)
Call TakeScreenShot
Call SetTableData("SAPLALDBINTERVAL", "Lower limit", 1, "", "", DT_S_ALR_87012357_3020_TABLECELL_LOWER_LIMIT_0,True)
Call SetTableData("SAPLALDBINTERVAL", "Upper limit", 1, "", "", DT_S_ALR_87012357_3020_TABLECELL_UPPER_LIMIT_0,True)
Call SetTableData("SAPLALDBINTERVAL", "Lower limit", 2, "", "", DT_S_ALR_87012357_3020_TABLECELL_LOWER_LIMIT_1,True)
Call SetTableData("SAPLALDBINTERVAL", "Upper limit", 2, "", "", DT_S_ALR_87012357_3020_TABLECELL_UPPER_LIMIT_1,True)
Call TakeScreenShot
Call ClickButton("Copy   \(F8\)", True)
Call ClickButton("Execute   \(F8\)", False)
Call PressEnter()

Call ClickButton("Local file...   \(Ctrl\+Shift\+F9\)", False)
Call SelectRadioButton("SPOPLI-SELFLAG", "Text with Tabs", True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)", True)

Call SetTextbox("Directory","DY_PATH","",DT_S_ALR_87012357_0200_DIRECTORY,True)
Call SetTextbox("File Name","DY_FILENAME","",DT_S_ALR_87012357_0200_FILE_NAME,True)
Call TakeScreenShot

Call ClickButton("Replace existing file   \(Ctrl\+S\)", True)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC34, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC1, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC2, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC3, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC4, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC5, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC6, 1)

Call ClickButton("Next page   \(Page down\)", False)

Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC7, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC8, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC9, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC10, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC11, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC12, 1)
Call  VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC13, 1)
Call ClickButton("Next page   \(Page down\)", False)


Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC14, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC15, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC16, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC17, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC18, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC19, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC19_1, 1)
Call ClickButton("Next page   \(Page down\)", False)

Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC20, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC21, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC22, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC23, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC24, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC25, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC26, 1)

Call ClickButton("Next page   \(Page down\)", False)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC27, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC28, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC35, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC29, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC30, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC31, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_TAX_CODE_OCC32, 1)
Call ClickButton("Next page   \(Page down\)", False)
Call TakeScreenShot


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
