'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.2.1.1.3. Create Vendor Account Group ZVXT -  External Trade Ve
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_2.2.1.1.3. Create Vendor Account Group ZVXT -  External Trade Ve"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_013_ENA_prices_are_not_higher_than_AB_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

''----------------------Tcode XK01----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


Call SetTextbox("Company Code","RF02K-BUKRS","",DT_XK01_0100_COMPANY_CODE,False)
Call SetTextbox("PurchasingOrganization","RF02K-EKORG","",DT_XK01_0100_PURCHASINGORGANIZATION,False)
Call SetTextbox("Account Group","RF02K-KTOKK","",DT_XK01_0100_ACCOUNT_GROUP,False)
Call TakeScreenShot()

Call PressEnter()  
Wait(2)

Call SetTextbox("Name","ADDR1_DATA-NAME1","",DT_XK01_0301_NAME,False)
Call SetTextbox("Search term 1/2","ADDR1_DATA-SORT1","","",False)
Call SetTextbox("Street/House number","ADDR1_DATA-STREET","",DT_XK01_0301_STREETHOUSE_NUMBER,False)
Call SetTextbox("Street/House number","ADDR1_DATA-HOUSE_NUM1","",DT_XK01_0301_STREETHOUSE_NUMBER_OCC1,False)
Call SetTextbox("Postal Code/City","ADDR1_DATA-POST_CODE1","",DT_XK01_0301_POSTAL_CODECITY,False)
Call SetTextbox("Postal Code/City","ADDR1_DATA-CITY1","",DT_XK01_0301_POSTAL_CODECITY_OCC1,False)
Call SetTextbox("Country","ADDR1_DATA-COUNTRY","",DT_XK01_0301_COUNTRY,False)
Call SetTextbox("Telephone","SZA1_D0100-TEL_NUMBER","",DT_XK01_0301_TELEPHONE,False)
Call SetTextbox("Fax","SZA1_D0100-FAX_NUMBER","",DT_XK01_0301_FAX,False)
Call SetTextbox("E-Mail","SZA1_D0100-SMTP_ADDR","",DT_XK01_0301_EMAIL,False)
Call SetTextbox("Region","ADDR1_DATA-REGION","",DT_XK01_0301_REGION,False)
Call SetTextbox("Mobile Phone","SZA1_D0100-MOB_NUMBER","",DT_XK01_0301_MOBILE_PHONE,False)
Call TakeScreenShot()

Call FocusTextBox("E-Mail","SZA1_D0100-SMTP_ADDR",False)
Call TakeScreenShot()

Call ClickButtonIfExist("Other E-Mail Addresses",False)
Wait(2)
Call TakeScreenShot()

Call SetTableDataNoRef("SAPLSZA6T_CONTROL6","Notes",1,DT_XK01_0600_TABLECELL_NOTES_0,True)
Call ClickButtonIfExist("Copy   \(Enter\)",True)
Wait(2)
Call TakeScreenShot()

Call SetTextbox("Data line","ADDR1_DATA-EXTENSION1","",DT_XK01_0301_DATA_LINE,False)
Call SetTextbox("Comments","ADDR1_DATA-REMARK","",DT_XK01_0301_COMMENTS,False)
Call TakeScreenShot()

Call ClickButtonIfExist("Next screen   \(F8\)",False)
Wait(2)
Call TakeScreenShot()


Call SelectCheckbox("LFA1-STKZU","1",DT_XK01_0120_SALESPURTAX,False)
Call SetTextbox("Location no\. 1","LFA1-BBBNR","",DT_XK01_0120_LOCATION_NO_1,False)
Call SetTextbox("Location no\. 2","LFA1-BBSNR","",DT_XK01_0120_LOCATION_NO_2,False)
Call SetTextbox("Check digit","LFA1-BUBKZ","",DT_XK01_0120_CHECK_DIGIT,False)
Call SetTextbox("Tax Number 1","LFA1-STCD1","",DT_XK01_0120_TAX_NUMBER_1,False)
Call SetTextbox("Tax Number 3","LFA1-STCD3","",DT_XK01_0120_TAX_NUMBER_3,False)
Call TakeScreenShot()
Call PressEnter()  

Call SetTableDataNoRef("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","Ctry",1,DT_XK01_0130_TABLECELL_CTRY_0,False)
Call SetTableDataNoRef("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","Bank Key",1,DT_XK01_0130_TABLECELL_BANK_KEY_0,False)
Call SetTableDataNoRef("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","Bank Account",1,DT_XK01_0130_TABLECELL_BANK_ACCOUNT_0,False)
Call SetTableDataNoRef("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","BnkT",1,DT_XK01_0130_TABLECELL_BNKT_0,False)
Call TakeScreenShot()
Call PressEnter()  

Call ClickButtonIfExist("Next screen   \(F8\)",False)
Wait(2)
Call ClickButtonIfExist("Next screen   \(F8\)",False)
Wait(2)

Call Sendkey("{F8}")

Call SetTextbox("Recon\. account","LFB1-AKONT","",DT_XK01_0210_RECON_ACCOUNT,False)
Call SetTextbox("Sort key","LFB1-ZUAWA","",DT_XK01_0210_SORT_KEY,False)

Call ClickButtonIfExist("Next screen   \(F8\)",False)
Wait(2)

Call SelectCheckbox("LFB1-REPRF","1",DT_XK01_0215_CHK_DOUBLE_INV,False)
Call SetTextbox("Pmnt Terms","LFB1-ZTERM","",DT_XK01_0215_PAYT_TERMS,False)
Call TakeScreenShot()

Call FocusTextBox("Payment methods","LFB1-ZWELS",False)

Call SendKey("{F4}")
wait 5
Call SetTableDataNoRef("SAPMF02KTCTRL_ZAHLWEGE","Indicator: Payment mtd selected(outgoing)?",4,"ON",False)

Call PressEnter()  


Call SetTextbox("Pmnt Meth\. Sup\.","LFB1-UZAWE","",DT_XK01_0215_PMT_METHSUPL,False)
Call TakeScreenShot

Call ClickButtonIfExist("Next screen   \(F8\)",False)
Call TakeScreenShot

Call SetTextbox("Acct statement","LFB1-XAUSZ","",DT_XK01_0220_ACCT_STATEMENT,False)
Call TakeScreenShot

Call ClickButtonIfExist("Next screen   \(F8\)",False)

Call SelectCheckbox("LFM1-PRFRE","1","ON",False)
Call SelectCheckbox("LFM1-NRGEW","1","ON",False)
Call SelectCheckbox("LFM1-BOLRE","1","ON",False)
Call SelectCheckbox("LFM1-WEBRE","1","ON",False)
Call SelectCheckbox("LFM1-AGREL","1","ON",False)

Call SetTextbox("Order currency","LFM1-WAERS","",DT_XK01_0310_ORDER_CURRENCY,False)
Call SetTextbox("Terms of paymnt","LFM1-ZTERM","",DT_XK01_0310_TERMS_OF_PAYMNT,False)
Call SetTextbox("Purchasing group","LFM1-EKGRP","",DT_XK01_0310_PURCHASING_GROUP,False)
Call SetTextbox("Planned deliv\. time","LFM1-PLIFZ","",DT_XK01_0310_PLANNED_DELIV_TIME,False)
Call TakeScreenShot

Call ClickButton("Sub-ranges   \(Shift\+F4\)",False)

Call SetTableDataNoRef("SAPMF02KTCTRL_TEILSORTIMENTE","Vendor Subrange",1,DT_XK01_0360_TABLECELL_VENDOR_SUBRANGE_0,False)
Call SetTableDataNoRef("SAPMF02KTCTRL_TEILSORTIMENTE","Vendor Subrange",2,DT_XK01_0360_TABLECELL_VENDOR_SUBRANGE_1,False)
Call SetTableDataNoRef("SAPMF02KTCTRL_TEILSORTIMENTE","VSR description",1,DT_XK01_0360_TABLECELL_VSR_DESCRIPTION_0,False)
Call SetTableDataNoRef("SAPMF02KTCTRL_TEILSORTIMENTE","VSR description",2,DT_XK01_0360_TABLECELL_VSR_DESCRIPTION_1,False)

Call ClickButton("Back   \(F3\)",False)

Call ClickButton("Different Data   \(Shift\+F6\)",False)
Call ClickButtonIfExist("SPOP-OPTION1",False)
Call TakeScreenShot

Call SelectCheckbox("WRF02K-HINZE","1","ON",False)
Call SetTextbox("Vendor Subrange","WRF02K-LTSNR","",DT_XK01_0352_VENDOR_SUBRANGE,True)
Call TakeScreenShot
Call PressEnter()
wait 5

Call SetTextbox("Terms of paymnt","LFM2-ZTERM","",DT_XK01_0330_TERMS_OF_PAYMNT,False)
Call SetTextbox("Planned deliv\. time","LFM2-PLIFZ","",DT_XK01_0330_PLANNED_DELIV_TIME,False)
Call TakeScreenShot
Call PressEnter()

Call ClickButton("Create   \(F5\)",False)
Call TakeScreenShot

Call SetTextbox("Vendor Subrange","WRF02K-LTSNR","",DT_XK01_0352_VENDOR_SUBRANGE_OCC1,True)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Terms of paymnt","LFM2-ZTERM","",DT_XK01_0330_TERMS_OF_PAYMNT_OCC1,False)
Call SetTextbox("Planned deliv\. time","LFM2-PLIFZ","",DT_XK01_0330_PLANNED_DELIV_TIME_OCC1,False)
Call TakeScreenShot
Call PressEnter()

Call ClickButton("Back   \(F3\)",False)
Call ClickButtonIfExist("Next screen   \(F8\)",False)
Call ClickButton("Save   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_XK01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_XK01_0100_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus ()
'''''
''Set objWsh = CreateObject("WScript.Shell") 
''objWsh.SendKeys "{F4}" 
''Set objWsh=nothing
'''''
''''''Capture the screenshot
'''''Call TakeScreenShot()
'''''
'''''Call PressEnter()  
''''''Capture the screenshot
'''''Call TakeScreenShot()
'''''
'''''Call SendKey("{F4}")
'''''Wait(1)
'''''Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)
''''''Capture the screenshot
'''''Call TakeScreenShot()
'''''
'''''Call ClickLabel("22711401","",True)
'''''Call ClickButtonIfExist("Copy   \(Enter\)",True)
''''''Capture the screenshot
'''''Call TakeScreenShot()
'''''
'''''Call FocusTextBox("Sort key","LFB1-ZUAWA",False)
'''''Call SendKey("{F4}")
'''''Wait(1)
'''''Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC3)
''''''Capture the screenshot
'''''Call TakeScreenShot()
'''''
'''''Call ClickLabel("001","",True)
'''''Call ClickButtonIfExist("Copy   \(Enter\)",True)
''''''Capture the screenshot
'''''Call TakeScreenShot()
'''''
'''''Call ClickButtonIfExist("Next screen   \(F8\)",False)
'''''Wait(2)
''''''Capture the screenshot
'''''Call TakeScreenShot()
'''''
'''''Call SelectCheckbox("LFB1-REPRF","1",DT_XK01_0215_CHK_DOUBLE_INV,False)
'''''Call SetTextbox("Payt Terms","LFB1-ZTERM","",DT_XK01_0215_PAYT_TERMS,False)
''''''Capture the screenshot
'''''Call TakeScreenShot()
'''''
'''''Call PressEnter()  
''''''Capture the screenshot
'''''Call TakeScreenShot()
'''''
'''''Call PressEnter()  
''''''Capture the screenshot
'''''Call TakeScreenShot()
'''''
'''''Call SetTextbox("Order currency","LFM1-WAERS","",DT_XK01_0310_ORDER_CURRENCY,False)
''''''Capture the screenshot
'''''Call TakeScreenShot()
'''''
'''''Call ClickButton("Save   \(Ctrl\+S\)",False)
'''''Wait(3)
''''''Validate If Customer is created
'''''
