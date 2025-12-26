'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_03-01-01-01-01-Create Vendor_V2 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_03-01-01-01-01-Vend_V2"
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

''--------TransactionCode-XK01----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE) 

Call SetTextbox("Company Code", "RF02K-BUKRS", "", DT_XK01_0100_COMPANY_CODE, False)
Call SetTextbox("PurchasingOrganization", "RF02K-EKORG", "", DT_XK01_0100_PURCHASINGORGANIZATION, False)
Call SetTextbox("Account Group", "RF02K-KTOKK", "", DT_XK01_0100_ACCOUNT_GROUP, False)
Call TakeScreenShot
Call PressEnter()
Call SetCombo("Title", DT_XK01_0301_TITLE)
Call SetTextbox("Name", "ADDR1_DATA-NAME1", "", DT_XK01_0301_NAME, False)
Call SetTextbox("Search term 1/2", "ADDR1_DATA-SORT1", "", DT_XK01_0301_SEARCH_TERM_12, False)
Call SetTextbox("Street/House number", "ADDR1_DATA-STREET", "", DT_XK01_0301_STREETHOUSE_NUMBER, False)
Call SetTextbox("Street/House number", "ADDR1_DATA-HOUSE_NUM1", "", DT_XK01_0301_STREETHOUSE_NUMBER_OCC1, False)
Call SetTextbox("Postal Code/City", "ADDR1_DATA-POST_CODE1", "", DT_XK01_0301_POSTAL_CODECITY, False)
Call SetTextbox("Postal Code/City", "ADDR1_DATA-CITY1", "", DT_XK01_0301_POSTAL_CODECITY_OCC1, False)
Call SetTextbox("Country", "ADDR1_DATA-COUNTRY", "", DT_XK01_0301_COUNTRY, False)
Call SetCombo("Language", DT_XK01_0301_LANGUAGE)
Call SetTextbox("Telephone", "SZA1_D0100-TEL_NUMBER", "", DT_XK01_0301_TELEPHONE, False)
Call SetTextbox("Mobile Phone", "SZA1_D0100-MOB_NUMBER", "",DT_XK01_0301_MOBILE_PHONE, False)
Call TakeScreenShot
Call ClickButton("Other E-Mail Addresses",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call SetTableDataNoRef("SAPLSZA6T_CONTROL6","E-Mail Address",1,DT_XK01_0600_TABLECELL_EMAIL_ADDRESS_0,True)
Call SetTableDataNoRef("SAPLSZA6T_CONTROL6","Notes",1,DT_XK01_0600_TABLECELL_NOTES_0,True)
Call ClickButton("New Line   \(Shift\+F1\)",True)
Call SetTableDataNoRef("SAPLSZA6T_CONTROL6","E-Mail Address",2,DT_XK01_0600_TABLECELL_EMAIL_ADDRESS_1,True)
Call SetTableDataNoRef("SAPLSZA6T_CONTROL6","Notes",2,DT_XK01_0600_TABLECELL_NOTES_1,True)
Call ClickButton("New Line   \(Shift\+F1\)",True)
Call SetTableDataNoRef("SAPLSZA6T_CONTROL6","E-Mail Address",3,DT_XK01_0600_TABLECELL_EMAIL_ADDRESS_2,True)
Call SetTableDataNoRef("SAPLSZA6T_CONTROL6","Notes",3,DT_XK01_0600_TABLECELL_NOTES_2,True)
Call TakeScreenShot
Call ClickButton("Copy   \(Enter\)", True)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("VAT Reg. No.", "LFA1-STCEG", "","", False)
Call TakeScreenShot
Call PressEnter() 'This step should be deleted once having Valid VAT Reg No.
Call PressEnter()
Call ClickCellTable("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","IBAN",1," "," ",False)
Call ClickButton("Switch Input Type", True)
Call SetTextbox("IBAN", "IBAN00", "", DT_XK01_0200_IBAN,True)
Call SetTextbox("SWIFT/BIC", "BNKA-SWIFT", "", DT_XK01_0200_SWIFT_CODE,True)
Call TakeScreenShot
Call ClickButton("Generate Bank Details   \(F7\)", True)
Call ClickButton("Continue   \(Enter\)", True)
Call SetTableDataNoRef("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","Acct Holder",1,DT_XK01_0130_TABLECELL_ACCT_HOLDER_0,False)
Call SetTableDataNoRef("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","BnkT",1,DT_XK01_0130_TABLECELL_BNKT_0,False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call SetTextbox("Recon. account", "LFB1-AKONT", "", DT_XK01_0210_RECON_ACCOUNT,False)
Call SetTextbox("Sort key", "LFB1-ZUAWA", "", DT_XK01_0210_SORT_KEY,False)
Call TakeScreenShot
Call PressEnter()
Call SelectCheckbox("LFB1-REPRF", 1,DT_XK01_0215_CHK_DOUBLE_INV, False)
Call SetTextboxNoLabel("LFB1-ZTERM", "", DT_XK01_0215_PAYT_TERMS,False)
Call SetTextbox("Cr memo terms", "LFB1-GUZTE", "", DT_XK01_0215_CR_MEMO_TERMS,False)
Call SetTextbox("Payment methods", "LFB1-ZWELS", "",DT_XK01_0215_PAYMENT_METHODS ,False)
Call SetTextbox("Tolerance group", "LFB1-TOGRR", "",DT_XK01_0215_TOLERANCE_GROUP_OCC1 ,False)
Call SetTextbox("Assign.Grp", "LFB1-ASSIGN_TEST", "",DT_XK01_0215_ASSIGNGRP ,False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call SelectCheckbox("LFM1-AGREL", 1,"ON", False)
Call SelectCheckbox("LFM1-BLIND", 1,DT_XK01_0310_DOC_INDEX_ACTIVE, False)
Call SetTextbox("Order currency", "LFM1-WAERS", "",DT_XK01_0310_ORDER_CURRENCY ,False)
Call SetTextbox("Terms of paymnt", "LFM1-ZTERM", "",DT_XK01_0310_TERMS_OF_PAYMNT ,False)
Call TakeScreenShot
Call SelectMenuBar("Environment;Classification")
Call SetTableDataNoRef("SAPLCLFMTC_OBJ_CLASS","Class",1,DT_XK01_1600_TABLECELL_CLASS_0,False)
Call PressEnter()
Call SetTableDataNoRef("SAPLCTMSCHARS_S","Value",5,DT_XK01_5100_TABLECELL_VALUE_4,False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Save   \(Ctrl\+S\)", False)
Call TakeScreenShot
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
Call GetStatusBar("item1", "DT_OP_XK01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR")
GetRowNo =6
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_XK01_0100_CHECK_TEXT_OF_STATUSBAR) 
Call TakeScreenShot

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
