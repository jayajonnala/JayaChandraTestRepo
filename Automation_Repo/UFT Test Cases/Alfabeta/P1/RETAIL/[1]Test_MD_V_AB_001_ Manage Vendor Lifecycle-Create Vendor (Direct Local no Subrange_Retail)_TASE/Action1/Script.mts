

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_V_AB_001_ Manage Vendor Lifecycle-Create Vendor (Direct Local no Subrange_Retail)
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


''gstrTestCaseName = "Test_MD_V_AB_001_ Manage Vendor Lifecycle-Create Vendor (Direct Local no Subrange_Retail)"
gstrTestCaseName = "Test_Create Vendor (Direct Local no Subrange_Retail)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\P1_FICO\DT_04.04.02.21 VIM - PO Precontrole Issue - BR01 - Invalid Company_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =3
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''
''''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''--------------------------------------------  XK01------------------------------------------------
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Account Group","RF02K-KTOKK","",DT_XK01_0100_ACCOUNT_GROUP,False)
Call SetTextbox("Company Code","RF02K-BUKRS","",DT_XK01_0100_COMPANY_CODE,False)
Call SetTextbox("PurchasingOrganization","RF02K-EKORG","",DT_XK01_0100_PURCHASINGORGANIZATION,False)
'Call SetTextbox("Vendor","RF02K-LIFNR",0,DT_XK01_0100_VENDOR,False)
Call TakeScreenShot()
Call PressEnter() 

Call SetTextbox("Name","ADDR1_DATA-NAME1","",DT_XK01_0301_NAME,False)
Call SetTextbox("Search term 1/2","ADDR1_DATA-SORT1","",DT_XK01_0301_SEARCH_TERM_12,False)
Call SetTextbox("Street/House number","ADDR1_DATA-STREET","",DT_XK01_0301_STREETHOUSE_NUMBER,False)
Call SetTextbox("Street/House number","ADDR1_DATA-HOUSE_NUM1","",DT_XK01_0301_STREETHOUSE_NUMBER_OCC1,False)
Call SetTextbox("Postal Code/City","ADDR1_DATA-POST_CODE1","",DT_XK01_0301_POSTAL_CODECITY,False)
Call SetTextbox("Postal Code/City","ADDR1_DATA-CITY1","",DT_XK01_0301_POSTAL_CODECITY_OCC1,False)
Call SetTextbox("Country","ADDR1_DATA-COUNTRY","",DT_XK01_0301_COUNTRY,False)
Call SetTextbox("Telephone","SZA1_D0100-TEL_NUMBER","",DT_XK01_0301_TELEPHONE,False)
Call SetTextbox("Mobile Phone","SZA1_D0100-MOB_NUMBER","",DT_XK01_0301_MOBILE_PHONE,False)
Call SetTextbox("Fax","SZA1_D0100-FAX_NUMBER","",DT_XK01_0301_FAX,False)
Call SetTextbox("E-Mail","SZA1_D0100-SMTP_ADDR","",DT_XK01_0301_EMAIL,False)
Call TakeScreenShot()
Call ClickButton("Other E-Mail Addresses",False)
Call SetTableData("SAPLSZA6T_CONTROL6","Notes",1,"","",DT_XK01_0600_TABLECELL_NOTES_0,True)
Call TakeScreenShot()
Call ClickButton("Copy   \(Enter\)",True)
Call ClickButton("International versions",False)
Call SelectRowGuiGridbyRowNo("",0,2,True)
Call DoubleClick()
Call SetCombo("SZA1_D0100-TITLE_MEDI",DT_XK01_0201_TITLE)
Call SetTextbox("Name","ADDR1_DATA-NAME1","",DT_XK01_0201_NAME,True)
Call SetTextbox("Search term 1/2","ADDR1_DATA-SORT1","",DT_XK01_0201_SEARCH_TERM_12,True)
Call SetTextbox("Street/House number","ADDR1_DATA-STREET","",DT_XK01_0201_STREETHOUSE_NUMBER,True)
Call SetTextbox("Street/House number","ADDR1_DATA-HOUSE_NUM1","",DT_XK01_0201_STREETHOUSE_NUMBER_OCC1,True)
Call SetTextbox("Postal Code/City","ADDR1_DATA-CITY1","",DT_XK01_0201_POSTAL_CODECITY,True)
Call TakeScreenShot()
Call ClickButton("Copy   \(Enter\)",True)
Call PressEnter() 
Call SetTextbox("VAT Reg\. No\.","LFA1-STCEG","",DT_XK01_0120_VAT_REG_NO,False) ' --- Getting error message as VAT Registration is already used by other Vendor .. Hence commented
Call SetTextbox("Tax Number 2","LFA1-STCD2","",DT_XK01_0120_TAX_NUMBER_2,False)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT3",DT_INCREMENT3+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''
''''Call SetSpecialTextbox("Type of Busines","LFA1-J_1KFTBUS",0,DT_XK01_TYPE_OF_BUSINESS,False)
Call FocusTextBox("Type of Busines","LFA1-J_1KFTBUS", False)
Call SendKey("{F4}")
Wait 10
Call ClickButton("Find   \(Ctrl\+F\)",True)
Call SetTextbox("Find","RSYSF-STRING","",DT_XK01_0800_FIND,True)
Call SelectCheckbox("SCAN_STRING-START",0,DT_XK01_0800_STARTING_AT_CURRENT_LINE,True)
Call TakeScreenShot()
Call ClickButton("Find   \(Enter\)",True)
Call SetFocusGuiLabel(DT_XK01_0800_FIND,"","",True)
Call  ClickButton("Continue   \(Enter\)",True)
Call  ClickButton("Copy   \(Enter\)",True)

Call SetTextbox("Tax office","LFA1-FISKU","",DT_XK01_0120_TAX_OFFICE,False)
Call SetTextbox("Location no\. 1","LFA1-BBBNR","",DT_XK01_0120_LOCATION_NO_1,False)
Call SetTextbox("Location no\. 2","LFA1-BBSNR","",DT_XK01_0120_LOCATION_NO_2,False)
Call SetTextbox("Check digit","LFA1-BUBKZ","",DT_XK01_0120_CHECK_DIGIT,False)
Call SetTextbox("Industry","LFA1-BRSCH","",DT_XK01_0120_INDUSTRY,False)
Call TakeScreenShot()

Call PressEnter()
Call PressEnter() 
Call PressEnter() 
Call PressEnter() 
Call PressEnter() 
Call PressEnter() 

Call SetTextbox("Recon\. account","LFB1-AKONT","",DT_XK01_0210_RECON_ACCOUNT,False)

Call SetTextbox("Sort key","LFB1-ZUAWA","",DT_XK01_0210_SORT_KEY,False)
Call SetTextbox("Prev\.acct no\.","LFB1-ALTKN","",DT_XK01_0210_PREVACCT_NO,False)
Call SetTextbox("Cash mgmnt group","LFB1-FDGRV","",DT_XK01_0210_CASH_MGMNT_GROUP,False)
Call TakeScreenShot()
Call PressEnter()

Call SelectCheckbox("LFB1-REPRF",0,DT_XK01_0215_CHK_DOUBLE_INV,False)
Call SetTextboxNoLabel("LFB1-ZTERM","",DT_XK01_0215_PAYT_TERMS,False)
Call SetTextbox("Cr memo terms","LFB1-GUZTE","",DT_XK01_0215_CR_MEMO_TERMS,False)
Call SetTextbox("Payment methods","LFB1-ZWELS","",DT_XK01_0215_PAYMENT_METHODS,False)
Call SetTextbox("Alternat\.payee","LFB1-LNRZB","",DT_XK01_0215_ALTERNATPAYEE,False)
Call TakeScreenShot()
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call PressEnter()

Call SelectCheckbox("LFM1-WEBRE",0,DT_XK01_0310_GRBASED_INV_VERIF,False)
Call SelectCheckbox("LFM1-BOLRE",0,DT_XK01_0310_SUBSEQUENT_SETTLEMENT,False)
Call SelectCheckbox("LFM1-BOIND",0,DT_XK01_0310_SUBSEQ_SETT_INDEX,False)
Call SelectCheckbox("LFM1-NRGEW",0,DT_XK01_0310_GRANT_DISCOUNT_IN_KIND,False)
Call SelectCheckbox("LFM1-PRFRE",0,DT_XK01_0310_RELEVANT_FOR_PRICE_DET__VENDOR_HIERARCHY,False)
Call SelectCheckbox("LFM1-AGREL",0,DT_XK01_0310_RELEVANT_FOR_AGENCY_BUSINESS,False)
Call SetTextbox("Order currency","LFM1-WAERS","",DT_XK01_0310_ORDER_CURRENCY,False)
Call SetTextbox("Terms of paymnt","LFM1-ZTERM","",DT_XK01_0310_TERMS_OF_PAYMNT,False)
Call SetTextbox("Pricing Date Control","LFM1-MEPRF","",DT_XK01_0310_PRICING_DATE_CONTROL,False)
Call SetTextbox("Purchasing group","LFM1-EKGRP","",DT_XK01_0310_PURCHASING_GROUP,False)
Call SetTextbox("Planned deliv\. time","LFM1-PLIFZ","",DT_XK01_0310_PLANNED_DELIV_TIME,False)
Call TakeScreenShot()
Call PressEnter()

Call SelectMenuBar("Environment;Classification")
Call SetTableData("SAPLCLFMTC_OBJ_CLASS","Class",1,"","",DT_XK01_1600_TABLECELL_CLASS_0,False)
Call PressEnter()
Call SetTableData("SAPLCTMSCHARS_S","Value",1,"","",DT_XK01_5100_TABLECELL_VALUE_0,False)
Call SetTableData("SAPLCTMSCHARS_S","Value",2,"","",DT_XK01_5100_TABLECELL_VALUE_1,False)
Call SetTableData("SAPLCTMSCHARS_S","Value",8,"","",DT_XK01_5100_TABLECELL_VALUE_7,False)
Call PressEnter()
Call TakeScreenShot()
Call SetTableData("SAPLCTMSCHARS_S","Value",22,"","",DT_XK01_5100_TABLECELL_VALUE_1_OCC1,False)
Call PressEnter()
Call TakeScreenShot()
Call SetTableData("SAPLCTMSCHARS_S","Value",27,"","","",False)
Call SetTableData("SAPLCTMSCHARS_S","Value",23,"","",DT_XK01_5100_TABLECELL_VALUE_2,False)
Call PressEnter()
Call TakeScreenShot()
Call SetTableData("SAPLCTMSCHARS_S","Value",24,"","",DT_XK01_5100_TABLECELL_VALUE_2_TRIANGULAR,False)
Call SetTableData("SAPLCTMSCHARS_S","Value",27,"","",DT_XK01_5100_TABLECELL_VALUE_2_DIRECTVENDORS,False)
Call PressEnter()
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Legacy Vendor ID .*",False)

Call SetTableData("SAPLZMDVM_LEGACY_VENDOR_ADDONSGTC_0100","Legacy Vendor Nr",1,"","",DT_XK01_0100_TABLECELL_LEGACY_VENDOR_NR_0,False)
Call PressEnter()
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_XK01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")

Call VerifyStatusBar("Vendor "&DT_XK01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" has been created for company code "& DT_XK01_0100_COMPANY_CODE &" purchasing organization "& DT_XK01_0100_PURCHASINGORGANIZATION)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_XK01_0100_CHECK_TEXT_OF_STATUSBAR)
'
Call LogOff()
Call FinalStatus ()




