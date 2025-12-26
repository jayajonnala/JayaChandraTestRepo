'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_LUXEMBOURG Vat Return
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_LUXEMBOURG Vat Return"
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
'
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
Call TakeScreenShot
Call ClickButton("%_BR_BELNR_%_APP_%-VALU_PUSH", False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 3, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_2,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 4, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_3,True)
Call SelectTab("TAB_STRIP", "Select Ranges", True)
Call TakeScreenShot
Call SetTableData("SAPLALDBINTERVAL", "Lower limit", 1, "", "", DT_S_ALR_87012357_3020_TABLECELL_LOWER_LIMIT_0,True)
Call SetTableData("SAPLALDBINTERVAL", "Lower limit", 2, "", "", DT_S_ALR_87012357_3020_TABLECELL_LOWER_LIMIT_1,True)
Call SetTableData("SAPLALDBINTERVAL", "Upper limit", 1, "", "", DT_S_ALR_87012357_3020_TABLECELL_UPPER_LIMIT_0,True)
Call SetTableData("SAPLALDBINTERVAL", "Upper limit", 2, "", "", DT_S_ALR_87012357_3020_TABLECELL_UPPER_LIMIT_1,True)
Call TakeScreenShot
Call ClickButton("Copy   \(F8\)", True)
Call TakeScreenShot
Call ClickButton("PUSHB_O1", False)
Call SelectCheckbox("PAR_XSTW",0,"ON", False)
Call SetTextbox("Tax Reporting Country","SEL_LSTM-LOW","",DT_S_ALR_87012357_1000_TAX_RETURN_COUNTRY,False)
Call TakeScreenShot
Call ClickButton("%_SEL_MWKZ_%_APP_%-VALU_PUSH", False)
Call SelectTab("TAB_STRIP", "Exclude Single Values", True)
Call SetTableData("SAPLALDBSINGLE_E", "Single value", 1, "", "", DT_S_ALR_87012357_3030_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE_E", "Single value", 2, "", "", DT_S_ALR_87012357_3030_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableData("SAPLALDBSINGLE_E", "Single value", 3, "", "", DT_S_ALR_87012357_3030_TABLECELL_SINGLE_VALUE_2,True)
Call SetTableData("SAPLALDBSINGLE_E", "Single value", 4, "", "", DT_S_ALR_87012357_3030_TABLECELL_SINGLE_VALUE_3,True)
Call TakeScreenShot
Call ClickButton("Copy   \(F8\)", True)
Call ClickButton("PUSHB_O2", False)
Call SetTextbox("Postings document type","PAR_BLAR","",DT_S_ALR_87012357_1000_POSTINGS_DOCUMENT_TYPE,False) 
Call SelectCheckbox("PAR_BINP",0,"ON", False)
Call SetTextbox("Posting date","PAR_BDAT","",ConvertDate(DT_S_ALR_87012357_1000_POSTING_DATE_OCC1),False)
Call SetTextbox("Due Date of the Tax Payable","PAR_FDAT","",ConvertDate(DT_S_ALR_87012357_1000_DUE_DATE_OF_THE_TAX_PAYABLE),False)
Call TakeScreenShot
Call ClickButton("PUSHB_O3", False)
Call SelectCheckbox("PAR_NAVA",0,"ON", False)
Call SelectCheckbox("PAR_XSHT",0,"ON", False)
Call SelectCheckbox("PAR_XADR",0,"ON", False)
Call TakeScreenShot
Call ClickButton("PUSHB_O4", False)
Call SelectCheckbox("PAR_LIS5",0,"ON", False)
Call TakeScreenShot
Call ClickButton("PUSHB_O5", False)
Call SelectCheckbox("PARPEUVA",0,"ON", False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)", False)
Call ClickButtonIfExist("with text was not displayed ", True)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
'Call ClickButton("Next page   \(Page down\)", False)
'Call ClickButton("Next page   \(Page down\)", False)
'Call ClickButton("Next page   \(Page down\)", False)
Call TakeScreenShot
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_NO_NAME, 1)

''''''--------TransactionCode--SM35---------''''

Call SetTcode(DT_S_ALR_87012357_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Created by:","D0100-CREATOR","",DT_S_ALR_87012357_1005_CREATED_BY,False)
Call TakeScreenShot
Call PressEnter()
Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1, False)
Call ClickButton("Process session   \(F8\)",False)
Call SelectRadioButton("D0300-ERROR", "Display errors only", True)
Call TakeScreenShot
Call ClickButton("Process   \(Enter\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Go back to batch input session overview   \(Enter\)",True)
Call TakeScreenShot

'''''--------TransactionCode-/nfotv----------''''

Call SetTcode(DT_S_ALR_87012357_1000_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Dominant Enterprise/Comp\. Code","SO_ORG-LOW","",DT_S_ALR_87012357_1000_DOMENTCOCODE,False)
Call SetTextbox("Return Type","PA_TYP","",DT_S_ALR_87012357_1000_RETURN_TYPE,False)
Call SetTextboxNoLabel("PA_CTRY","",DT_S_ALR_87012357_1000_REPORTING_COUNTRY,False)
Call SetTextbox("ReportingPeriod","SO_PER-LOW","",DT_S_ALR_87012357_1000_REPORTINGPERIOD,False)
Call SetTextbox("Rep. Year","PA_RJAHR","",DT_S_ALR_87012357_1000_REP_YEAR,False)
Call PressEnter()
Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot    
Call ClickButton("Refresh   \(Ctrl\+F2\)", False)
Call ClickButton("Display Details on Package   \(Shift\+F7\)", False)
Call TakeScreenShot
Call ClickButton("Local File...   \(F6\)", False)
Call SelectRadioButton("SPOPLI-SELFLAG", "Text with Tabs", True)
Call ClickButton("Continue   \(Enter\)", True)
Call SetTextbox("Directory","DY_PATH","",DT_S_ALR_87012357_0200_DIRECTORY,True)
Call SetTextbox("File Name","DY_FILENAME","",DT_S_ALR_87012357_0200_FILE_NAME,True)
Call TakeScreenShot
Call ClickButton("Replace existing file   \(Ctrl\+S\)", True)
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM,0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC1, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC2, 0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC3, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC4, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC5, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC6, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC7, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC8, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC9, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC10, 0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC11, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC12, 0)
Call ClickButton("Last Page   \(Ctrl\+Page down\)", False)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC13, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC14, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC15, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC16, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC17, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC18, 0)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC19, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC20, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC21, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC22, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC23, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC24, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC25, 0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0200_200_CHECK_TEXT_OF_PARAM_OCC26, 0)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
