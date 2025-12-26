'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_E10G07P02S01V01 Rep on Gen Tax Docs from BELUX for BE10
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_E10G07P02S01V01 from BELUX for BE10"
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

''''''--------TransactionCode-S_ALR_87012357----------''''

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
Call ClickButton("%_BR_BELNR_%_APP_%-VALU_PUSH", False)
Call SelectTab("TAB_STRIP", "Select Ranges", True)
Call TakeScreenShot
Call SetTableData("SAPLALDBINTERVAL", "Lower limit", 1, "", "", DT_S_ALR_87012357_3020_TABLECELL_LOWER_LIMIT_0,True)
Call SetTableData("SAPLALDBINTERVAL", "Lower limit", 2, "", "", DT_S_ALR_87012357_3020_TABLECELL_LOWER_LIMIT_1,True)
Call SetTableData("SAPLALDBINTERVAL", "Lower limit", 3, "", "", DT_S_ALR_87012357_3020_TABLECELL_LOWER_LIMIT_2,True)
Call SetTableData("SAPLALDBINTERVAL", "Upper limit", 1, "", "", DT_S_ALR_87012357_3020_TABLECELL_UPPER_LIMIT_0,True)
Call SetTableData("SAPLALDBINTERVAL", "Upper limit", 2, "", "", DT_S_ALR_87012357_3020_TABLECELL_UPPER_LIMIT_1,True)
Call SetTableData("SAPLALDBINTERVAL", "Upper limit", 3, "", "", DT_S_ALR_87012357_3020_TABLECELL_UPPER_LIMIT_2,True)
Call TakeScreenShot
Call SelectTab("TAB_STRIP", "Select Single Values", True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 3, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_2,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 4, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_3,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 5, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_4,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 6, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_5,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 7, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_6,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 8, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_7,True)
Call TakeScreenShot
Call ClickButton("Copy   \(F8\)", True)
Call ClickButton("Execute   \(F8\)", False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Local file...   \(Ctrl\+Shift\+F9\)", False)
Call SelectRadioButton("SPOPLI-SELFLAG", "Text with Tabs", True)
Call ClickButton("Continue   \(Enter\)", True)
Call SetTextbox("Directory","DY_PATH","",DT_PATH,True)
Call SetTextbox("File Name","DY_FILENAME","",DT_FILENAME,True)
Call TakeScreenShot
Call ClickButton("Replace existing file   \(Ctrl\+S\)", True)
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_NO_NAME,0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_NO_NAME_OCC4,0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_1A,0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_1B,0)
Call ClickButton("Next page   \(Page down\)", False)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_1C,0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_1D,0)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_C1,0)
'Call ClickButton("Next page   \(Page down\)", False)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_C2,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_C3,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_D1,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_D2,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_D3,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_DC,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_K1,0)
'Call ClickButton("Next page   \(Page down\)", False)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_K2,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_K3,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_NB,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_P3,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_R3,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_S1,0)
'Call ClickButton("Next page   \(Page down\)", False)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_U0,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_U1,0)
''Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_N4,0)
''Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_N5,0)
''Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_N6,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_A1,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_A2,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_A3,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_B0,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_B2,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_B3,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_C1,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_C2,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_C3,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_C8,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_D1,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_D2,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_D3,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_DC,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_DJ,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_J3,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_K1,1)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_K2,1)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_K3,1)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_N1,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_N2,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_N3,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_N7,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_ND,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_NE,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_NG,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_O1,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_P3,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_R3,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_S1,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_S3,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_X2,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_U0,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_U1,0)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
