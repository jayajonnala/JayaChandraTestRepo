'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_AUTOMATED IC LISTING BELGIUM
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_AUTOMATED IC LISTING BELGIUM"
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

''''Call CloseSessionsSAP()
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
Call ClickButton("PUSHB_O2", False)
Call SetTextbox("Postings document type","PAR_BLAR","",DT_S_ALR_87012357_1000_POSTINGS_DOCUMENT_TYPE,False) 
Call TakeScreenShot
Call ClickButton("PUSHB_O3", False) 
Call SelectCheckbox("PAR_NAVA",0,"ON", False)
Call SelectCheckbox("PAR_XSHT",0,"ON", False)
Call SelectCheckbox("PAR_XADR",0,"ON", False)
Call ClickButton("PUSHB_O4", False)
Call SelectCheckbox("PAR_LIS5",0,"ON", False)
Call TakeScreenShot
Call ClickButton("PUSHB_O5", False)
Call SelectCheckbox("PARPEUVA",0,"ON", False)
Call TakeScreenShot
Call ClickButton("%_BR_BELNR_%_APP_%-VALU_PUSH", False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SelectTab("TAB_STRIP", "Select Ranges", True)
Call TakeScreenShot
Call SetTableData("SAPLALDBINTERVAL", "Lower limit", 1, "", "", DT_S_ALR_87012357_3020_TABLECELL_LOWER_LIMIT_0,True)
Call SetTableData("SAPLALDBINTERVAL", "Upper limit", 1, "", "", DT_S_ALR_87012357_3020_TABLECELL_UPPER_LIMIT_0,True)
Call ClickButton("Copy   \(F8\)", True)
Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot
Call ClickButtonIfExist("with text was not displayed ", True)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
Call ClickButton("Next page   \(Page down\)", False)
'Call ClickButton("Next page   \(Page down\)", False)
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_NO_NAME, 1)

'''''--------TransactionCode-/nfotv----------''''

Call SetTcode(DT_S_ALR_87012357_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot

'Call SetTextbox("Dom.Ent./CoCode","SO_ORG-LOW","",DT_S_ALR_87012357_1000_DOMENTCOCODE,False)
' SetTextboxNoLabel(textboxName, textboxIndex, textboxValue, blnIsItPopup)
Call SetTextboxNoLabel("SO_ORG-LOW","",DT_S_ALR_87012357_1000_DOMENTCOCODE,False)
Call SetTextbox("Return Type","PA_TYP","",DT_S_ALR_87012357_1000_RETURN_TYPE,False)
'Call SetTextbox("Reporting Country","PA_CTRY","",DT_S_ALR_87012357_1000_REPORTING_COUNTRY,False)
Call SetTextboxNoLabel("PA_CTRY","",DT_S_ALR_87012357_1000_REPORTING_COUNTRY,False)
Call SetTextbox("ReportingPeriod","SO_PER-LOW","",DT_S_ALR_87012357_1000_REPORTINGPERIOD,False)
Call SetTextbox("Rep. Year","PA_RJAHR","",DT_S_ALR_87012357_1000_REP_YEAR,False)
Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot    
Call ClickButton("Refresh   \(Ctrl\+F2\)", False)
Call ClickButton("Display Details on Package   \(Shift\+F7\)", False)
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_NO_NAME_OCC3, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_NO_NAME_OCC4, 1)
Call VerifyifGuiLabelExists(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_NO_NAME_OCC1)
Call VerifyifGuiLabelExists(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_NO_NAME_OCC2)

'''--------TransactionCode-/ns_alr_87012402----------''''

Call SetTcode(DT_S_ALR_87012357_0120_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot

Call SelectCheckbox("PAR_BSET",0,"ON", False)
Call SelectCheckbox("PAR_SRV",0,"ON", False)
Call SelectCheckbox("PAR_EPOS",0,"ON", False)
Call TakeScreenShot
Call SetTextbox("Telephone no.of contact person","PAR_TELF","",DT_S_ALR_87012357_1000_TELEPHONE_NOOF_CONTACT_PERSON,False)
Call SetTextbox("File name in file system","PAR_UNIX","",DT_S_ALR_87012357_1000_FILE_NAME_IN_FILE_SYSTEM,False)
Call SetTextbox("Contact person","PAR_CONT","",DT_S_ALR_87012357_1000_CONTACT_PERSON,False)
Call SetTextbox("Company Code","BR_BUKRS-LOW","",DT_S_ALR_87012357_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","",DT_S_ALR_87012357_1000_FISCAL_YEAR_OCC1,False)
Call SetTextbox("Posting date","BR_BUDAT-LOW","",ConvertDate(DT_S_ALR_87012357_1000_POSTING_DATE_OCC1),False)
Call SetTextbox("to","BR_BUDAT-HIGH","",ConvertDate(DT_S_ALR_87012357_1000_TO_OCC1),False)
Call SetTextbox("Reporting country","SEL_LSTM-LOW","",DT_S_ALR_87012357_1000_REPORTING_COUNTRY_OCC1,False)
Call SetTextbox("Output Tax Code","SEL_MWKZ-LOW","",DT_S_ALR_87012357_1000_OUTPUT_TAX_CODE,False)
Call SetTextbox("Period for Posting Date Selec.","P_MONAPD","",DT_S_ALR_87012357_1000_PERIOD_FOR_POSTING_DATE_SELEC,False)
Call SetTextbox("/","P_JAMOPD","","2024",False)
Call SetTextbox("Sending Company's VAT Number","P_VATNUM","",DT_S_ALR_87012357_0200_SENDING_COMPANYS_VAT_NUMBER,False)
Call SetTextbox("Sending Company's Name","P_NAME","",DT_S_ALR_87012357_0200_SENDING_COMPANYS_NAME,False)
Call SetTextbox("Street","P_STREET","",DT_S_ALR_87012357_0200_STREET,False)
Call SetTextbox("Post Code","P_POST","",DT_S_ALR_87012357_0200_POST_CODE,False)
Call SetTextbox("City","P_CITY","",DT_S_ALR_87012357_0200_CITY,False)
Call SetTextbox("Phone","P_BETEL","",DT_S_ALR_87012357_0200_PHONE,False)
Call SetTextbox("Email Address","P_EMAIL","",DT_S_ALR_87012357_0200_EMAIL_ADDRESS,False)
Call SetTextbox("Country","P_CNTRY","",DT_S_ALR_87012357_0200_COUNTRY,False)
Call TakeScreenShot
Call ClickButton("%_BR_BELNR_%_APP_%-VALU_PUSH", False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_0_OCC1,True)
Call SelectTab("TAB_STRIP", "Select Ranges", True)
Call TakeScreenShot
Call SetTableData("SAPLALDBINTERVAL", "Lower limit", 1, "", "", DT_S_ALR_87012357_3020_TABLECELL_LOWER_LIMIT_0_OCC1,True)
Call SetTableData("SAPLALDBINTERVAL", "Upper limit", 1, "", "", DT_S_ALR_87012357_3020_TABLECELL_UPPER_LIMIT_0_OCC1,True)
Call ClickButton("Copy   \(F8\)", True)
Call ClickButton("%_SEL_MWKZ_%_APP_%-VALU_PUSH", False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_0_OCC2,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 3, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_2,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 4, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_3,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 5, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_4,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 6, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_5,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 7, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_6,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 8, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_7,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 9, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_6_OCC1,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 10, "", "", DT_S_ALR_87012357_3010_TABLECELL_SINGLE_VALUE_7_OCC1,True)
Call TakeScreenShot
Call ClickButton("Copy   \(F8\)", True)
Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot
Call ClickButton("Next page   \(Page down\)", False)
'Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_NO_NAME_OCC5, 1)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
