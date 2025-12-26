
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AR0035 Clear AR Accounts (Manual & Automatic) - automatic clearing in EUR for Customers (without special GL)_TASE
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


gstrTestCaseName = "Test_AR0035_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AR0001 Manage AR Documents-Manage Manual Customer Invoicing_Cred.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''----------------------Tcode F-22----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",Replace((DT_F22_100_DOCUMENT_DATE),"/","."),False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace((DT_F22_100_POSTING_DATE),"/","."),False)
'Call SetTextbox("Type","BKPF-BLART","",DT_OB08_0100_TYPE,False)
'Call SetTextbox("Period","BKPF-MONAT","",DT_F22_100_PERIOD,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F22_100_COMPANY_CODE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F22_100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F22_100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F22_100_DOCHEADER_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_100_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

'Call SelectCheckbox("BKPF-XMWST","1",DT_OB08_0301_CALCULATE_TAX,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_301_AMOUNT,False)
Call SetTextbox("Bus\. Area","BSEG-GSBER","",DT_F22_301_BUS_AREA,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_301_ASSIGNMENT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_301_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_301_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_301_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

Call ClickButton("Previous item   \(Shift\+F6\)",False)
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next
Call GetTextboxValue("BSEG-DMBTR","","DT_AMOUNT_LC_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_301_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_301_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F22_300_TAX_CODE,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_300_ASSIGNMENT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_300_TEXT,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F22_1006_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F22_1006_COST_CENTER,False)
'Call SetTextbox("Profit Center","COBL-PRCTR","",DT_F22_1006_PROFIT_CENTER,False)
'Capture the screenshot
Call TakeScreenShot()

'Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
''Capture the screenshot
'Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

'veryfy sattus bar content
Call GetStatusBar("item1","DT_F22_100_GET_TEXT_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'VerifyStatusBar(DT_F22_100_CHECK_TEXT_OF_STATUSBAR)
'
''----------------------Tcode FB03----------------------------

'Enter the Tcode
Call SetTcode(DT_F22_100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_F22_100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F22_100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_F22_1000_FISCAL_YEAR,False)

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Change Display Currency   \(F8\)",False)
'SAPGuiSession("Session").SAPGuiWindow("Display Currency for Document").InsightObject("InsightObject").Click
Call ClickButton("Continue/Confirm   \(Enter\)",True)
Call VerifyGridCellContent("",1,"Amount","",DT_F22_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR)
Call VerifyGridCellContent("",2,"Amount","",DT_F22_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMBTR)

Wait(1)
Call ClickButton("Exit   \(Shift\+F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

'''----------------------Tcode F-22----------------------------
''Enter the Tcode
Call SetTcode(DT_F22_100_OKCD_OCC2) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",Replace((DT_F22_100_DOCUMENT_DATE_OCC2),"/","."),False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace((DT_F22_100_POSTING_DATE_OCC2),"/","."),False)
'Call SetTextbox("Type","BKPF-BLART","",DT_OB08_0100_TYPE,False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F22_100_PERIOD,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F22_100_COMPANY_CODE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F22_100_CURRENCYRATE_OCC2,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F22_100_REFERENCE_OCC2,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F22_100_DOCHEADER_TEXT_OCC2,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_100_PSTKY_OCC2,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_100_ACCOUNT_OCC2,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

'Call SelectCheckbox("BKPF-XMWST","1",DT_OB08_0301_CALCULATE_TAX,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_301_AMOUNT_OCC2,False)
Call SetTextbox("Bus\. Area","BSEG-GSBER","",DT_F22_301_BUS_AREA_OCC2,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_301_ASSIGNMENT_OCC2,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_301_TEXT_OCC2,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_301_PSTKY_OCC2,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_301_ACCOUNT_OCC2,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

Call ClickButton("Previous item   \(Shift\+F6\)",False)
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next
Call GetTextboxValue("BSEG-DMBTR","","DT_AMOUNT2_LC_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_301_PSTKY_OCC2,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_301_ACCOUNT_OCC2,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_301_AMOUNT_OCC3,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F22_300_TAX_CODE_OCC2,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_300_ASSIGNMENT_OCC2,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_300_TEXT_OCC2,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F22_1006_BUSINESS_AREA_OCC2,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F22_1006_COST_CENTER_OCC2,False)
'Call SetTextbox("Profit Center","COBL-PRCTR","",DT_F22_1006_PROFIT_CENTER,False)
'Capture the screenshot
'Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
''Capture the screenshot
'Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next
Wait(2)
'veryfy sattus bar content
Call GetStatusBar("item1","DT_F22_100_GET_TEXT_OF_STATUSBAR_OCC2_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'VerifyStatusBar(DT_F22_100_CHECK_TEXT_OF_STATUSBAR)

''----------------------Tcode FB03----------------------------

'Enter the Tcode
Call SetTcode(DT_F22_100_OKCD_OCC3) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_F22_100_DOCUMENT_NUMBER_OCC2,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F22_100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_F22_1000_FISCAL_YEAR,False)

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call VerifyGridCellContent("",1,"BSCHL","",DT_F22_750_CHECK_POSTING_KEY_1)
Call VerifyGridCellContent("",2,"BSCHL","",DT_F22_750_CHECK_POSTING_KEY_2)
Call VerifyGridCellContent("",1,"KTONR","",DT_F22_750_CHECK_ACCOUNT_1)
Call VerifyGridCellContent("",2,"KTONR","",DT_F22_750_CHECK_ACCOUNT_2)
Call ClickButton("Change Display Currency   \(F8\)",False)
'SAPGuiSession("Session").SAPGuiWindow("Display Currency for Document").InsightObject("InsightObject").Click
Call ClickButton("Continue/Confirm   \(Enter\)",True)
Call VerifyGridCellContent("",2,"Amount","",DT_F22_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMBTR_OCC2)
Call VerifyGridCellContent("",1,"Amount","",DT_F22_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR_OCC2)
'Call VerifyGridCellContent("",2,"Amount","",DT_F22_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMBTR_OCC2)
'Call VerifyGridCellContent("",1,"Amount","",DT_F22_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR_OCC2)

Wait(1)
Call ClickButton("Exit   \(Shift\+F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)

''----------------------Tcode FBL5N----------------------------
'Enter the Tcode
Call SetTcode(DT_F22_100_OKCD_OCC4) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC7)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_F22_1000_CUSTOMER_ACCOUNT,False)
Call SelectRadioButton("X_AISEL","All items",False)
Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_F22_100_COMPANY_CODE,False)

Call ActivateNodeGuiTree("","Customer master;Industry")
Call ActivateNodeGuiTree("","Customer master;Group key")
Call ActivateNodeGuiTree("","Customer master;Country")
Call ActivateNodeGuiTree("","Customer master;Postal Code")
Call ActivateNodeGuiTree("","Customer master;City")
Call ActivateNodeGuiTree("","Customer master;Trading partner")
Call ActivateNodeGuiTree("","Documents;Special G/L ind.")
Call ActivateNodeGuiTree("","Documents;Assignment")
Call ActivateNodeGuiTree("","Company code;Clerk Abbreviation")
Call ActivateNodeGuiTree("","Company code;Reconciliation acct")
'Capture the screenshot
Call TakeScreenShot()

wait(1)

Call ClickButtonByIndex("Multiple selection",4,False)
'SapGuiSession("transaction:=FBL5N").SapGuiWindow("transaction:=FBL5N").SAPGuiButton("tooltip:=Multiple selection","index:=4").Click

'Capture the screenshot
Call TakeScreenShot()

'set filter criteria
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_F22_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_F22_3010_TABLECELL_SINGLE_VALUE_1,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Copy   \(F8\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

'verify the details
Call VerifyifGuiLabelExists_ByIndex(DT_F22_120_CHECK_TEXT_OF_NO_NAME,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_120_CHECK_TEXT_OF_NO_NAME_OCC4,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_120_CHECK_TEXT_OF_NO_NAME_OCC2,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_120_CHECK_TEXT_OF_NO_NAME_OCC5,0)
'Call VerifyGridCellContent("",1,"Document Number","",DT_F22_120_CHECK_TEXT_OF_NO_NAME)
'Call VerifyGridCellContent("",2,"Document Number","",DT_F22_120_CHECK_TEXT_OF_NO_NAME_OCC4)

'Call VerifyGridCellContent("",1,"Document Date","",DT_F22_120_CHECK_TEXT_OF_NO_NAME_OCC2)
'Call VerifyGridCellContent("",2,"Document Date","",DT_F22_120_CHECK_TEXT_OF_NO_NAME_OCC5)

Call VerifyifGuiLabelExists_ByIndex(DT_F22_120_CHECK_TEXT_OF_NO_NAME_OCC3,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_120_CHECK_TEXT_OF_NO_NAME_OCC6,0)
Call VerifyifGuiLabelExists_ByIndex(DT_CLEARANCE_AMOUNT,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_F22_120_CHECK_TEXT_OF_NO_NAME_OCC5,0)
'Call VerifyGridCellContent("",1,"Amount in local currency","",DT_F22_120_CHECK_TEXT_OF_NO_NAME_OCC3)
'Call VerifyGridCellContent("",2,"Amount in local currency","",DT_F22_120_CHECK_TEXT_OF_NO_NAME_OCC6)
'Call VerifyGridCellContent("",3,"Amount in local currency","",DT_CLEARANCE_AMOUNT)

Wait(1)
Call ClickButton("Exit   \(Shift\+F3\)",False)
Call ClickButton("Exit   \(Shift\+F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC8)

''----------------------Tcode F13E----------------------------

'Enter the Tcode
Call SetTcode(DT_F22_100_OKCD_OCC5) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC9)

Call SelectCheckbox("XTOL","1",DT_F22_1000_INCLUDE_TOLERANCES,False)
Call SelectCheckbox("X_KUNNR","1",DT_F22_1000_SELECT_CUSTOMERS,False)
Call SetTextbox("Company Code","BUKRX-LOW","",DT_F22_1000_COMPANY_CODE,False)
'Call SetTextbox("to","BUKRX-HIGH","","",False)'DT_F22_1000_TO
Call SetTextbox("Fiscal Year","GJAHX-LOW","",DT_F22_1000_FISCAL_YEAR,False)
Call SetTextbox("Document Number","DOCNR-LOW","",DT_F22_3010_TABLECELL_SINGLE_VALUE_1_OCC2,False)
Call SetTextbox("to","DOCNR-HIGH","",DT_F22_3010_TABLECELL_SINGLE_VALUE_0_OCC2,False)
Call SetTextbox("Customers","KONTD-LOW","",DT_F22_1000_CUSTOMERS,False)
Call SetTextbox("Clearing Currency","ZWAERS","",DT_F22_1000_CLEARING_CURRENCY,False)
''Capture the screenshot
'Call TakeScreenShot()
'
'Call ClickButton("Execute   \(F8\)",False)
''Capture the screenshot
'Call TakeScreenShot()
'
'Call ClickButton("Back   \(F3\)",False)
Call SelectCheckbox("X_TESTL","1",DT_F22_1000_TEST_RUN,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F13E").SapGuiWindow("transaction:=F13E").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next
'Capture the screenshot
Call TakeScreenShot()

'veryfy sattus bar content
Call GetStatusBar("item1","DT_F22_120_GET_TEXT_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'VerifyStatusBar(DT_F22_120_CHECK_TEXT_OF_STATUSBAR)

Call VerifyifGuiLabelExists(DT_F22_120_CHECK_TEXT_OF_NO_NAME_OCC7)
Call VerifyifGuiLabelExists(DT_F22_120_CHECK_TEXT_OF_NO_NAME_OCC8)

Wait(1)
Call ClickButton("Exit   \(Shift\+F3\)",False)
Call ClickButton("Exit   \(Shift\+F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC10)
'
'''----------------------Tcode FB03----------------------------
'
'Enter the Tcode
Call SetTcode(DT_F22_100_OKCD_OCC6) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC11)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_F22_100_DOCUMENT_NUMBER_OCC3,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F22_100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_F22_1000_FISCAL_YEAR,False)

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Display Document Header   \(F5\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call VerifyTextBoxContent("Document type","BKPF-BLART","",DT_FB03_1710_CHECK_TEXT_OF_DOCUMENT_TYPE,False)
Call ClickButton("Continue/Confirm   \(Enter\)",True)

Call VerifyGridCellContent("",1,"Account","",DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"Account","",DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)

Call VerifyGridCellContent("",1,"AZBET","",DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("",2,"AZBET","",DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_AZBET)

Call ClickButton("Change Display Currency   \(F8\)",False)
'SAPGuiSession("Session").SAPGuiWindow("Display Currency for Document").InsightObject("InsightObject").Click
Call ClickButton("Continue/Confirm   \(Enter\)",True)
Call VerifyGridCellContent("",2,"Amount","",DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_DMBTR)
Call VerifyGridCellContent("",1,"Amount","",DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR)
'Call VerifyGridCellContent("",1,"DMBTR","",DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR)
'Call VerifyGridCellContent("",2,"DMBTR","",DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_DMBTR)

Call VerifyGridCellContent("",1,"Posting Key","",DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"Posting Key","",DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)

Call VerifyGridCellContent("",1,"KOBEZ","",Lcase(DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOBEZ))
'
'''----------------------Tcode FBL5N----------------------------
'Enter the Tcode
Call SetTcode(DT_F22_100_OKCD_OCC7) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC13)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_F22_1000_CUSTOMER_ACCOUNT,False)

Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_F22_100_COMPANY_CODE,False)

Call ActivateNodeGuiTree("","Customer master;Industry")
Call ActivateNodeGuiTree("","Customer master;Group key")
Call ActivateNodeGuiTree("","Customer master;Country")
Call ActivateNodeGuiTree("","Customer master;Postal Code")
Call ActivateNodeGuiTree("","Customer master;City")
Call ActivateNodeGuiTree("","Customer master;Trading partner")
Call ActivateNodeGuiTree("","Documents;Special G/L ind.")
Call ActivateNodeGuiTree("","Documents;Assignment")
Call ActivateNodeGuiTree("","Company code;Clerk Abbreviation")
Call ActivateNodeGuiTree("","Company code;Reconciliation acct")
'Capture the screenshot
Call TakeScreenShot()

wait(1)
Call ClickButtonByIndex("Multiple selection",4,False)
'SapGuiSession("transaction:=FBL5N").SapGuiWindow("transaction:=FBL5N").SAPGuiButton("tooltip:=Multiple selection","index:=4").Click

'Capture the screenshot
Call TakeScreenShot()

'set filter criteria
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_F22_3010_TABLECELL_SINGLE_VALUE_0_OCC3,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_F22_3010_TABLECELL_SINGLE_VALUE_1_OCC3,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",3,DT_F22_3010_TABLECELL_SINGLE_VALUE_2,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Copy   \(F8\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRadioButton("X_CLSEL","Cleared items",False)
'Call SetTextbox("Clearing date","SO_AUGDT-LOW","","",False)'DT_F22_1000_CLEARING_DATE

Call ClickButton("Execute   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

Call VerifyifGuiLabelExists_ByIndex(DT_F22_120_CHECK_TEXT_OF_CL,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_120_CHECK_TEXT_OF_CL_OCC2,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_120_CHECK_TEXT_OF_CL_OCC2,0)
'Call VerifyGridCellContentByRefColumn("","","Document Number",DT_F22_3010_TABLECELL_SINGLE_VALUE_2,"Document type","",DT_F22_120_CHECK_TEXT_OF_CL)
'Call VerifyGridCellContentByRefColumn("","","Document Number",DT_F22_3010_TABLECELL_SINGLE_VALUE_0_OCC3,"Document type","",DT_F22_120_CHECK_TEXT_OF_CL_OCC2)

Wait(1)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()






'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [8,3437477]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//




