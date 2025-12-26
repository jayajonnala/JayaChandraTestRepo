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

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'reload DS to update dates and calculations
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_2.3.1.1.2. Create Customer with Account Group ZCPA - DG Public A_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 18th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_2.3.1.1.2. Create Customer with Account Group ZCPA - DG Public A_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_2.3.1.1.2. Create Customer with Account Group ZCPA - DG Public A_TASE.xls"
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

Call SetComboByKey("RF02D-KTOKD",DT_XD01_7100_ACCOUNT_GROUP)
Call SetTextbox("Company code","RF02D-BUKRS","",DT_XD01_7100_COMPANY_CODE,True)
Call SetTextbox("Sales Organization","RF02D-VKORG","",DT_XD01_7100_SALES_ORGANIZATION,True)
Call SetTextbox("Distribution Channel","RF02D-VTWEG","",DT_XD01_7100_DISTRIBUTION_CHANNEL,True)
Call SetTextbox("Division","RF02D-SPART","",DT_XD01_7100_DIVISION,True)
Call SetComboByKey("Account group",DT_XD01_7100_ACCOUNT_GROUP)
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(2)
Call TakeScreenShot()

Call SetTextbox("Name","ADDR1_DATA-NAME1","",DT_XD01_0301_NAME,False)
Call SetTextbox("Street/House number","ADDR1_DATA-STREET","",DT_XD01_0301_STREETHOUSE_NUMBER,False)
Call SetTextbox("Street/House number","ADDR1_DATA-HOUSE_NUM1","",DT_XD01_0301_STREETHOUSE_NUMBER_OCC1,False)
Call SetTextbox("Postal Code/City","ADDR1_DATA-POST_CODE1","",DT_XD01_0301_POSTAL_CODECITY,False)
Call SetTextbox("Postal Code/City","ADDR1_DATA-CITY1","",DT_XD01_0301_POSTAL_CODECITY_OCC1,False)
Call SetTextbox("Country","ADDR1_DATA-COUNTRY","",DT_XD01_0301_COUNTRY,False)
Call SetTextbox("Region","ADDR1_DATA-REGION","",DT_XD01_0301_REGION,False)
Call SetComboByKey("ADDR1_DATA-LANGU",DT_LANGUAGE)
Call SetTextbox("Telephone","SZA1_D0100-TEL_NUMBER","",DT_XD01_0301_TELEPHONE,False)
Call SetTextbox("Fax","SZA1_D0100-FAX_NUMBER","",DT_XD01_0301_FAX,False)
Call SetTextbox("E-Mail","SZA1_D0100-SMTP_ADDR","",DT_XD01_0301_EMAIL,False)
Call SetTextbox("Data line","ADDR1_DATA-EXTENSION1","",DT_XD01_0301_DATA_LINE,False)
Call SetTextbox("Comments","ADDR1_DATA-REMARK","",DT_XD01_0301_COMMENTS,False)
Call SetTextbox("Data line","ADDR1_DATA-EXTENSION1","",DT_XD01_0301_DATA_LINE,False)
Call TakeScreenShot()

Call FocusTextBox("Comments","ADDR1_DATA-REMARK",False)
Call TakeScreenShot()
Call SelectTab("TABSTRIP100","Control Data",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call TakeScreenShot()

Call SelectCheckbox("KNA1-STKZU","1",DT_XD01_7122_SALESPURTAX,False)

Call SetTextbox("Location no\. 1","KNA1-BBBNR","",DT_XD01_7123_LOCATION_NO_1,False)
Call SetTextbox("Location no\. 2","KNA1-BBSNR","",DT_XD01_7123_LOCATION_NO_2,False)
Call SetTextbox("Check digit","KNA1-BUBKZ","",DT_XD01_7123_CHECK_DIGIT,False)
Call SetTextbox("Tax Number 1","KNA1-STCD1","",DT_XD01_7122_TAX_NUMBER_1,False)
Call SetTextbox("Tax Number 3","KNA1-STCD3","",DT_XD01_7122_TAX_NUMBER_3,False)
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Payment Transactions",False)
Call TakeScreenShot()

Call SetTableDataNoRef("SAPMF02DTCTRL_ZAHLUNGSVERKEHR","Ctry",1,DT_XD01_7131_TABLECELL_CTRY_0,False)
Call SetTableDataNoRef("SAPMF02DTCTRL_ZAHLUNGSVERKEHR","Bank Key",1,DT_XD01_7131_TABLECELL_BANK_KEY_0,False)
Call SetTableDataNoRef("SAPMF02DTCTRL_ZAHLUNGSVERKEHR","Bank Account",1,DT_XD01_7131_TABLECELL_BANK_ACCOUNT_0,False)
Call SetTableDataNoRef("SAPMF02DTCTRL_ZAHLUNGSVERKEHR","Bk.typ.",1,DT_XD01_7131_TABLECELL_BKTYP_0,False)
Call TakeScreenShot()

Call ClickButtonIfExist("Company Code Data   \(Ctrl\+F2\)",False)
Wait(2)
Call TakeScreenShot()

Call SendKey("{F4}")
Wait(1)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
Call TakeScreenShot()

Call ClickLabel(DT_XD01_LABEL_1,"",True)
Call ClickButtonIfExist("Copy   \(Enter\)",True)
Call TakeScreenShot()

Call FocusTextBox("Sort key","KNB1-ZUAWA",False)
Call SendKey("{F4}")
Wait(1)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)
Call TakeScreenShot()

Call ClickLabel(DT_XD01_LABEL_2,"",True)
Call ClickButtonIfExist("Copy   \(Enter\)",True)
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Payment Transactions",False)
Call SetTextbox("Terms of payment","KNB1-ZTERM","",DT_XD01_7215_TERMS_OF_PAYMENT,False)
Call TakeScreenShot()

Call FocusTextBox("Payment methods","KNB1-ZWELS",False)
Call SendKey("{F4}")
Wait(1)
Call SetTableDataNoRef("SAPMF02DTCTRL_ZAHLWEGE","Indicator: Payment mtd selected(outgoing)?",2,DT_XD01_1215_TABLECELL_INDICATOR_PAYMENT_MTD_SELECTED_OUTGOING__1,True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot()

Call SetTextbox("Pmnt Meth\. Sup\.","KNB1-UZAWE","",DT_XD01_7216_PMT_METHSUPL,False)
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Correspondence",False)
Call TakeScreenShot()

Call ClickButtonIfExist("Sales Area Data   \(Ctrl\+F3\)",False)
Call TakeScreenShot()

Call SetTextbox("Cust\.pric\.proc\.","KNVV-KALKS","",DT_XD01_7311_CUSTPRICPROC,False)
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Shipping",False)
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Billing Documents",False)

Call SetTextbox("Terms of payment","KNVV-ZTERM","",DT_XD01_7321_TERMS_OF_PAYMENT,False)
Call SetTextbox("Acct assgmt group","KNVV-KTGRD","",DT_XD01_7322_ACCT_ASSGMT_GROUP,False)
Call SetTableDataNoRef("SAPMF02DTCTRL_STEUERN","Tax classification",1,DT_XD01_7350_TABLECELL_TAX_CLASSIFICATION_0,False)
Call TakeScreenShot()

Call PressEnter() 


Call SelectTab("TABSTRIP100","Partner Functions",False)
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(3)
Call GetStatusBar("item1","DT_GETMESSAGE_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_EXPECTED)

Call ClickButtonIfExist("Cancel   \(F12\)",True)

Call LogOff()
Call FinalStatus ()

