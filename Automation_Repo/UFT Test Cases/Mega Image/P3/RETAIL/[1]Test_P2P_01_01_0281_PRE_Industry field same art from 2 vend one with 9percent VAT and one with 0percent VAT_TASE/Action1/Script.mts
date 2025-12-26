
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0281_PRE_Industry field same art from 2 vend one with 9percent VAT and one with 0percent VAT
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
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)



gstrTestCaseName = "Test_P2P_01_01_0281_one with 0percent VAT"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_01-Regular purchasing in RW04  dry goods  via ME21N - P&Z_P3.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


''''''--------------login----------------'''''

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''--------TransactionCode-XK03----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'''-------------Select all checkboxes----------''''

Call SelectCheckbox("RF02K-D0110", 0, DT_XK03_0101_ADDRESS, False)
Call SelectCheckbox("RF02K-D0120", 1, DT_XK03_0101_CONTROL, False)
Call SelectCheckbox("RF02K-D0130", 2, DT_XK03_0101_PAYMENT_TRANSACTIONS, False)
Call SelectCheckbox("WRF02K-D0380", 3, DT_XK03_0101_CONTACT_PERSONS, False)
Call SelectCheckbox("RF02K-D0210", 4, DT_XK03_0101_ACCOUNTING_INFO, False)
Call SelectCheckbox("RF02K-D0215", 5, DT_XK03_0101_PAYMENT_TRANSACTIONS_OCC1, False)
Call SelectCheckbox("RF02K-D0220", 6, DT_XK03_0101_CORRESPONDENCE, False)
Call SelectCheckbox("RF02K-D0610", 7, DT_XK03_0101_WITHHOLDING_TAX, False)
Call SelectCheckbox("RF02K-D0310", 8, DT_XK03_0101_PURCHASING_DATA, False)
Call SelectCheckbox("WRF02K-D0320", 9, DT_XK03_0101_PARTNER_FUNCTIONS, False)


Call SetTextboxNoLabel("RF02K-LIFNR","",DT_XK03_0101_VENDOR,false)
Call SetTextbox("Company Code","RF02K-BUKRS","",DT_XK03_0101_COMPANY_CODE,false)
Call SetTextbox("Purchasing Organization","RF02K-EKORG","",DT_XK03_0101_PURCH_ORGANIZATION,false)
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Next screen   \(F8\)",False)
''Call VerifyTextBoxContent("Industry", "LFA1-BRSCH", "","" , False)
Call VerifyTextBoxContent("Industry", "LFA1-BRSCH", "",DT_XK03_0120_CHECK_TEXT_OF_INDUSTRY_OCC1 , False)
Call TakeScreenShot

Call ClickButton("Back   \(F3\)",False)
Call SetTextboxNoLabel("RF02K-LIFNR","",DT_XK03_0101_VENDOR_OCC1,false)
Call PressEnter()
Call ClickButton("Next screen   \(F8\)",False)
Call TakeScreenShot
''Call VerifyTextBoxContent("Industry", "LFA1-BRSCH", "", DT_XK03_0120_CHECK_TEXT_OF_INDUSTRY_OCC1, False)
Call VerifyTextBoxContent("Industry", "LFA1-BRSCH", "", "", False)

'''--------TransactionCode-MM43----------''''

Call SetTcode(DT_XK03_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_XK03_0120_OKCD)

Call SetTextbox("Article","RMMW1-MATNR","",DT_XK03_0100_ARTICLE,false)
Call SetTextboxNoLabel("RMMW1-LIFNR","","",false)
Call PressEnter()
Call TakeScreenShot
Call VerifyTextBoxContent("Tax class\.", "MARA-TAKLV", "", DT_XK03_2002_CHECK_TEXT_OF_TAX_CLASS, False)

Call ClickButton("Back   \(F3\)",False)

Call SetTextbox("Article","RMMW1-MATNR","",DT_XK03_0100_ARTICLE_OCC1,false)
Call SetTextbox("Purchasing Org\.","RMMW1-EKORG","",DT_XK03_0100_PURCHASING_ORG,false)
Call SetTextboxNoLabel("RMMW1-LIFNR","",DT_XK03_0100_VENDOR_OCC1,false)

Call ClickButton("Deselect All   \(Shift\+F7\)",False)
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100", "Screen description", DT_XK03_0100_SAPLMGMWTAB_CONT_0100_OCC1, false)
Call PressEnter()
Call TakeScreenShot

Call GetTextboxValue("EINE-NETPR", "", "DT_XK03_2223_CHECK_TEXT_OF_NET_PRICE", False)
Call VerifyTextBoxContent("Net Price", "EINE-NETPR", "", DT_XK03_2223_CHECK_TEXT_OF_NET_PRICE, False)
Call VerifyTextBoxContent("Net Price", "EINE-WAERS", "", DT_XK03_2223_CHECK_TEXT_OF_NET_PRICE_OCC2, False)

Call ClickButton("Back   \(F3\)",False)


Call SetTextbox("Article","RMMW1-MATNR","",DT_XK03_0100_ARTICLE_OCC1,false)
Call SetTextbox("Purchasing Org\.","RMMW1-EKORG","",DT_XK03_0100_PURCHASING_ORG,false)
Call SetTextboxNOLabel("RMMW1-LIFNR","",DT_XK03_0100_VENDOR_OCC2,false)

Call ClickButton("Deselect All   \(Shift\+F7\)",False)
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100", "Screen description", DT_XK03_0100_SAPLMGMWTAB_CONT_0100_OCC3, false)
Call PressEnter()
Call TakeScreenShot

Call GetTextboxValue("EINE-NETPR", "", "DT_XK03_2223_CHECK_TEXT_OF_NET_PRICE_OCC3", False)
Call VerifyTextBoxContent("Net Price", "EINE-NETPR", "", DT_XK03_2223_CHECK_TEXT_OF_NET_PRICE_OCC3, False)
Call VerifyTextBoxContent("Net Price", "EINE-WAERS", "", DT_XK03_2223_CHECK_TEXT_OF_NET_PRICE_OCC5, False)

Call LogOff()
Call FinalStatus ()

