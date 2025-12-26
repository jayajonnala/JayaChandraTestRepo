
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : 
'.................Test Scenario: Test_03-01-01-02-01-Modify Vendor General Data in Retail
'.................TCode: XK01
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_03-01-01-02-01- GD"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'ensure no open session
Call CloseSessionsSAP()
'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''--------TransactionCode-XK01----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE) 

Call SetTextbox("Company Code", "RF02K-BUKRS", "", DT_XK01_0100_COMPANY_CODE, False)
Call SetTextbox("PurchasingOrganization", "RF02K-EKORG", "", DT_XK01_0100_PURCHASINGORGANIZATION, False)
Call SetTextbox("Account Group", "RF02K-KTOKK", "", DT_XK01_0100_ACCOUNT_GROUP, False)
Call TakeScreenShot
Call PressEnter()
Call SetComboByKey("Title", DT_XK01_0301_TITLE)
Call SetTextbox("Name", "ADDR1_DATA-NAME1", "", DT_XK01_0301_NAME, False)
Call SetTextbox("Search term 1/2", "ADDR1_DATA-SORT1", "", DT_XK01_0301_SEARCH_TERM_12, False)
Call SetTextbox("Street/House number", "ADDR1_DATA-STREET", "", DT_XK01_0301_STREETHOUSE_NUMBER, False)
Call SetTextbox("Street/House number", "ADDR1_DATA-HOUSE_NUM1", "", DT_XK01_0301_STREETHOUSE_NUMBER_OCC1, False)
Call SetTextbox("Postal Code/City", "ADDR1_DATA-POST_CODE1", "", DT_XK01_0301_POSTAL_CODECITY, False)
Call SetTextbox("Postal Code/City", "ADDR1_DATA-CITY1", "", DT_XK01_0301_POSTAL_CODECITY_OCC1, False)
Call SetTextbox("Country", "ADDR1_DATA-COUNTRY", "", DT_XK01_0301_COUNTRY, False)
Call SetComboByKey("ADDR1_DATA-LANGU", "NL")
Call SetTextbox("Telephone", "SZA1_D0100-TEL_NUMBER", "", DT_XK01_0301_TELEPHONE, False)
Call SetTextbox("Fax", "SZA1_D0100-FAX_NUMBER", "", DT_XK01_0301_FAX, False)
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
Call SetComboByKey("Standard Method","INT")
Call TakeScreenShot
Call PressEnter()
'Call SetTextbox("VAT Reg. No.", "LFA1-STCEG", "",DT_XK01_0120_VAT_REG_NO, False)
'Call TakeScreenShot
Call PressEnter()
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
Call SetTextboxNoLabel( "LFB1-ZTERM", "", DT_XK01_0215_PAYT_TERMS,False)
''This function SetTextboxNoLabel is used for Pmnt and Payt label change.
Call SetTextbox("Cr memo terms", "LFB1-GUZTE", "", DT_XK01_0215_CR_MEMO_TERMS,False)
Call SetTextbox("Tolerance group", "LFB1-TOGRU", "", DT_XK01_0215_TOLERANCE_GROUP,False)
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
Call GetStatusBar("item1", "DT_OP_XK01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR")
GetRowNo =4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_XK01_0100_CHECK_TEXT_OF_STATUSBAR)
Call LogOff()
Call FinalStatus ()

'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [13,4062483]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

