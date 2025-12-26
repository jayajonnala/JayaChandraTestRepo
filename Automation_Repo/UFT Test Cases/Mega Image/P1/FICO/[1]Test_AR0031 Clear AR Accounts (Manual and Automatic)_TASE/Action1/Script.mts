'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_AR0031 Clear AR Accounts (Manual and Automatic)_TASE
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


gstrTestCaseName = "Test_AR0031 Clear AR Accounts (Manual and Automatic)_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''----------------------Tcode F-22----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",Replace((DT_F22_0100_DOCUMENT_DATE),"/","."),False)
'Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace((DT_F22_100_POSTING_DATE),"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F22_0100_TYPE,False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F22_0100_PERIOD,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F22_0100_COMPANY_CODE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F22_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F22_0100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F22_0100_DOCHEADER_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_0100_ACCOUNT,False)
'Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_F22_0100_SGL_IND,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

Call SelectCheckbox("BKPF-XMWST","1",DT_F22_0301_CALCULATE_TAX,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_0301_AMOUNT,False)
'SetTextboxNoLabel(textboxName, textboxIndex, textboxValue, blnIsItPopup)
Call SetTextboxNoLabel("BSEG-ZTERM","",DT_F22_0301_PAYT_TERMS,False)
'Call SetTextbox("Payt Terms","BSEG-ZTERM","",DT_F22_0301_PAYT_TERMS,False)
'Call SetTextbox("Pmnt Terms","BSEG-ZTERM","",DT_F22_0301_PAYT_TERMS,False)
'Call SetTextbox("Due on","BSEG-ZFBDT","",Replace((DT_F22_0303_DUE_ON),"/","."),False)
'Call SetTextbox("Bus\. Area","BSEG-GSBER","",DT_F22_301_BUS_AREA,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_0301_ASSIGNMENT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_0301_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_0301_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_0301_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F22_0300_TAX_CODE,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_0300_ASSIGNMENT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_0300_TEXT,False)
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
Call GetStatusBar("item1","DT_DOC_NO_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_F22_0100_CHECK_TEXT_OF_STATUSBAR)
'
'''----------------------Tcode F-22----------------------------
'
Call SetTextbox("Document Date","BKPF-BLDAT","",Replace((DT_F22_0100_DOCUMENT_DATE_OCC1),"/","."),False)
'Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace((DT_F22_100_POSTING_DATE),"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F22_0100_TYPE_OCC1,False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F22_0100_PERIOD_OCC1,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F22_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F22_0100_CURRENCYRATE_OCC1,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F22_0100_REFERENCE_OCC1,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F22_0100_DOCHEADER_TEXT_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_0100_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_0100_ACCOUNT_OCC1,False)
'Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_F22_0100_SGL_IND_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

Call SelectCheckbox("BKPF-XMWST","1",DT_F22_0301_CALCULATE_TAX_OCC1,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_0301_AMOUNT_OCC1,False)
'Call SetTextbox("Due on","BSEG-ZFBDT","",Replace((DT_F22_0303_DUE_ON_OCC1),"/","."),False)
'Call SetTextbox("Bus\. Area","BSEG-GSBER","",DT_F22_301_BUS_AREA,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_0301_ASSIGNMENT_OCC1,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_0301_TEXT_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_0301_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_0301_ACCOUNT_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_0300_AMOUNT_OCC1,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F22_0300_TAX_CODE_OCC1,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_0300_ASSIGNMENT_OCC1,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_0300_TEXT_OCC1,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F22_1006_BUSINESS_AREA_OCC1,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F22_1006_COST_CENTER_OCC1,False)
Call SetTextbox("Profit Center","COBL-PRCTR","",DT_F22_1006_PROFIT_CENTER,False)
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
Call GetStatusBar("item1","DT_DOC_NO_2_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_F22_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)

Wait(1)
Call ClickButton("Back   \(F3\)",False)
Call ClickButtonIfExist("Yes",True)
'Capture the screenshot
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
'
'''----------------------Tcode F-32----------------------------
'
'Enter the Tcode
Call SetTcode(DT_F22_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRadioButton("RF05A-XPOS1",DT_F22_0131_DOCUMENT_NUMBER,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F22_0131_ACCOUNT,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F22_0131_COMPANY_CODE,False)
'Call SetTextbox("Currency","BKPF-WAERS","",DT_OB08_0131_CURRENCY,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("From","RF05A-SEL01","0",DT_F22_0731_FROM,False)
Call SetTextbox("From","RF05A-SEL01","1",DT_F22_0731_FROM_OCC1,False)

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS","Partial Pmt",False)
Call SetTableDataNoRef("SAPDF05XTC_6104","Payment Amount",2,DT_F22_6104_TABLECELL_PAYMENT_AMOUNT_1,False)
Call PressEnter() 
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-32").SapGuiWindow("transaction:=F-32").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next 

Call ClickButton("Post   \(Ctrl\+S\)",False)
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-32").SapGuiWindow("transaction:=F-32").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next 

Wait(2)
'Validate If invoice is generated
Call GetStatusBar("item1","DT_DOC_NO_3_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_F22_0131_CHECK_TEXT_OF_STATUSBAR)

Wait(1)
Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
'
'''----------------------Tcode FB03----------------------------
'
'Enter the Tcode
Call SetTcode(DT_F22_0100_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_F22_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F22_0100_COMPANY_CODE_OCC2,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_F22_0100_FISCAL_YEAR,False)

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call VerifyGridCellContent("",1,"BSCHL","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"BSCHL","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
'Call VerifyGridCellContent("",3,"BSCHL","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)

Call VerifyGridCellContent("",1,"KTONR","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"KTONR","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
'Call VerifyGridCellContent("",3,"KTONR","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)

Call VerifyGridCellContent("",1,"AZBET","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("",2,"AZBET","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
'Call VerifyGridCellContent("",3,"AZBET","",DT_F22_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)

Wait(1)
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)
'
'''----------------------Tcode FBL5N----------------------------
''Enter the Tcode
Call SetTcode(DT_F22_0100_OKCD_OCC2) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_F22_1000_CUSTOMER_ACCOUNT,False)

Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_F22_1000_COMPANY_CODE,False)

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
SapGuiSession("transaction:=FBL5N").SapGuiWindow("transaction:=FBL5N").SAPGuiButton("tooltip:=Multiple selection","index:=4").Click

'Capture the screenshot
Call TakeScreenShot()

'set filter criteria
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_F22_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_F22_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",3,DT_F22_3010_TABLECELL_SINGLE_VALUE_2,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Copy   \(F8\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRadioButton("X_AISEL","All items",False)

Call ClickButton("Execute   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

''''''''''''Additional code to filter out the correct details as layout got changed to "grid type fro label format'''''''''''
'Call ClickButton("Change layout\.\.\.   \(Ctrl\+F8\)",False)
'Call SelectRowGuiGridbyRowNo("Column Set","",24,True)
'Call ClickButton("Show Selected Fields \(F7\)",True)
'Call ClickButton("Transfer   \(Enter\)",True)
'Call TakeScreenShot()

'verify the details
Call VerifyifGuiLabelExists_ByIndex("S_LEDR",0)
Call VerifyifGuiLabelExists_ByIndex("S_LEDR",1)
Call VerifyifGuiLabelExists_ByIndex("S_LEDG",0)
Call VerifyifGuiLabelExists_ByIndex("S_LEDG",1)
'Call VerifyGridCellContent("",1,"Cleared/open items symbol","",DT_F22_0120_CHECK_TOOLTIP_OF_NO_NAME)
'Call VerifyGridCellContent("",2,"Cleared/open items symbol","",DT_F22_0120_CHECK_TOOLTIP_OF_NO_NAME_OCC1)
'Call VerifyGridCellContent("",4,"Cleared/open items symbol","",DT_F22_0120_CHECK_TOOLTIP_OF_NO_NAME_OCC2)
'Call VerifyGridCellContent("",5,"Cleared/open items symbol","",DT_F22_0120_CHECK_TOOLTIP_OF_NO_NAME_OCC3)

Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_NO_NAME,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC1,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC3,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC4,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC5,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC2,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_RON,0)

'Call VerifyifGuiLabelExists_ByIndex(DT_F22_0120_CHECK_TEXT_OF_ACCOUNT_11003979,0)
'Call VerifyGridCellContent("",1,"Amount in local currency","",DT_F22_0120_CHECK_TEXT_OF_NO_NAME)
'Call VerifyGridCellContent("",2,"Amount in local currency","",DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC1)
'Call VerifyGridCellContent("",3,"Amount in local currency","",DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC2)
'Call VerifyGridCellContent("",5,"Amount in local currency","",DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC3)
'Call VerifyGridCellContent("",4,"Amount in local currency","",DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC4)
'Call VerifyGridCellContent("",7,"Amount in local currency","",DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC5)

'Call VerifyGridCellContent("",1,"Local Currency","",DT_F22_0120_CHECK_TEXT_OF_RON)
'Call VerifyGridCellContent("",1,"Account","",DT_F22_0120_CHECK_TEXT_OF_ACCOUNT_11003979)
Wait(1)


'***********************************************************************************************************************************************
'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()
'***********************************************************************************************************************************************
