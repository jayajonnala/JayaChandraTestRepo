'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'reload DS to update dates and calculations
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_2.2.1.1.2. Create Vendor Account Group ZVPA - Public Administra_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 18th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_2.2.1.1.2. Create Vendor Account Group ZVPA - Public Administra_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_2.2.1.1.2. Create Vendor Account Group ZVPA -  Public Administra_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode XD01----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call SetTextbox("Company Code","RF02K-BUKRS","",DT_XK01_0100_COMPANY_CODE,False)
Call SetTextbox("PurchasingOrganization","RF02K-EKORG","",DT_XK01_0100_PURCHASINGORGANIZATION,False)
Call SetTextbox("Account Group","RF02K-KTOKK","",DT_XK01_0100_ACCOUNT_GROUP,False)
Call TakeScreenShot()

Call PressEnter()  
Wait(2)
Call TakeScreenShot()

Call SetTextbox("Name","ADDR1_DATA-NAME1","",DT_XK01_0301_NAME,False)
Call SetTextbox("Search term 1/2","ADDR1_DATA-SORT1","",DT_XK01_0301_SEARCH_TERM_12,False)
Call SetTextbox("Street/House number","ADDR1_DATA-STREET","",DT_XK01_0301_STREETHOUSE_NUMBER,False)
Call SetTextbox("Street/House number","ADDR1_DATA-HOUSE_NUM1","",DT_XK01_0301_STREETHOUSE_NUMBER_OCC1,False)
Call SetTextbox("Postal Code/City","ADDR1_DATA-POST_CODE1","",DT_XK01_0301_POSTAL_CODECITY,False)
Call SetTextbox("Postal Code/City","ADDR1_DATA-CITY1","",DT_XK01_0301_POSTAL_CODECITY_OCC1,False)
Call SetTextbox("Country","ADDR1_DATA-COUNTRY","",DT_XK01_0301_COUNTRY,False)
Call SetTextbox("Telephone","SZA1_D0100-TEL_NUMBER","","",False)
Call SetTextbox("Fax","SZA1_D0100-FAX_NUMBER","",DT_XK01_0301_FAX,False)
Call SetTextbox("E-Mail","SZA1_D0100-SMTP_ADDR","",DT_XK01_0301_EMAIL,False)

Call TakeScreenShot()

Call FocusTextBox("E-Mail","SZA1_D0100-SMTP_ADDR",False)
Call TakeScreenShot()

Call ClickButtonIfExist("Other E-Mail Addresses",False)
Wait(2)
Call TakeScreenShot()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

Call SetTableDataNoRef("SAPLSZA6T_CONTROL6","Notes",1,DT_XK01_0600_TABLECELL_NOTES_0,True)
Call ClickButtonIfExist("Copy   \(Enter\)",True)
Wait(2)
Call TakeScreenShot()

Call SetTextbox("Comments","ADDR1_DATA-REMARK","",DT_XK01_0301_COMMENTS,False)
Call TakeScreenShot()

Call ClickButtonIfExist("Next screen   \(F8\)",False)
Wait(2)
Call TakeScreenShot()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)

Call SelectCheckbox("LFA1-STKZU","1",DT_XK01_0120_SALESPURTAX,False)
Call SetTextbox("Location no\. 1","LFA1-BBBNR","",DT_XK01_0120_LOCATION_NO_1,False)
Call SetTextbox("Location no\. 2","LFA1-BBSNR","",DT_XK01_0120_LOCATION_NO_2,False)
Call SetTextbox("Check digit","LFA1-BUBKZ","",DT_XK01_0120_CHECK_DIGIT,False)
Call SetTextbox("Tax Number 1","LFA1-STCD1","",DT_XK01_0120_TAX_NUMBER_1,False)
Call SetTextbox("Tax Number 3","LFA1-STCD3","",DT_XK01_0120_TAX_NUMBER_3,False)
Call TakeScreenShot()

Call PressEnter()  
Call TakeScreenShot()

Call SetTableDataNoRef("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","Ctry",1,DT_XK01_0130_TABLECELL_CTRY_0,False)
Call SetTableDataNoRef("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","Bank Key",1,DT_XK01_0130_TABLECELL_BANK_KEY_0,False)
Call SetTableDataNoRef("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","Bank Account",1,DT_XK01_0130_TABLECELL_BANK_ACCOUNT_0,False)
Call SetTableDataNoRef("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","BnkT",1,DT_XK01_0130_TABLECELL_BNKT_0,False)
Call TakeScreenShot()

Call PressEnter()  
Call TakeScreenShot()

Call PressEnter()  
Call TakeScreenShot()

Call PressEnter()  
Call TakeScreenShot()

Call SendKey("{F4}")
Wait(1)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)
Call TakeScreenShot()

Call ClickLabel(DT_XK01_0210_RECON_ACCOUNT,"",True)
Call SendKey("{F2}")
Wait(1)
Call TakeScreenShot()

Call FocusTextBox("Sort key","LFB1-ZUAWA",False)
Call SendKey("{F4}")
Wait(1)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC3)
Call TakeScreenShot()

Call ClickLabel(DT_XK01_0210_SORT_KEY,"",True)
Call SendKey("{F2}")
Wait(1)
Call TakeScreenShot()

Call ClickButtonIfExist("Next screen   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

Call SelectCheckbox("LFB1-REPRF","1",DT_XK01_0215_CHK_DOUBLE_INV,False)
Call SetTextbox("Pmnt Terms","LFB1-ZTERM","",DT_XK01_0215_PAYT_TERMS,False)
Call TakeScreenShot()

Call PressEnter()  
Call TakeScreenShot()

Call PressEnter()  
Call TakeScreenShot()

Call SetTextbox("Order currency","LFM1-WAERS","",DT_XK01_0310_ORDER_CURRENCY,False)
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(3)
Call GetStatusBar("item1","DT_XK01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_XK01_0100_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC4)

Call LogOff()
Call FinalStatus ()

