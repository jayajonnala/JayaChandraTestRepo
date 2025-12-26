		
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_GL0034 Clear GL Accounts (Manual & Automatic) - automatic cleari_TASE
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


gstrTestCaseName = "Test_GL0034"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call StartDateof445PeriodByDate(DT_TODAY,"DT_STARTING_DATE_PERIOD")
'

'''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''----------------------Tcode OB08----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call ClickButton("Position\.\.\.",False)
Call TakeScreenShot

'Call SetTextbox("Exch\. Rate Type","SVALD-VALUE","",DT_OB08_0300_EXCH_RATE_TYPE,False)
Call SetTextboxNoLabel("SVALD-VALUE","",DT_OB08_0300_EXCH_RATE_TYPE,False)

Call SetTextbox("From currency","SVALD-VALUE","",DT_OB08_0300_FROM_CURRENCY,False)
Call SetTextbox("To-currency","SVALD-VALUE","",DT_OB08_0300_TOCURRENCY,False)
Call SetTextbox("Valid from","SVALD-VALUE","",DT_OB08_0300_VALID_FROM,False)

Call ClickButton("Continue   \(Enter\)",True)     
'Call TakeScreenShot

Call GetTableCellData("SAPL0SAPTCTRL_V_TCURR", "Dir.quot.", 1, "", "", "DT_OB08_0020_CHECK_TEXT_OF_TABLECELL_DIRQUOT_0_OUTPUT", False)

Call ClickButton("Position\.\.\.",False)
Call TakeScreenShot

'Call SetTextbox("Exch\. Rate Type","SVALD-VALUE","",DT_OB08_0300_EXCH_RATE_TYPE_OCC1,False)
Call SetTextboxNoLabel("SVALD-VALUE","",DT_OB08_0300_EXCH_RATE_TYPE_OCC1,False)

Call SetTextbox("From currency","SVALD-VALUE","",DT_OB08_0300_FROM_CURRENCY_OCC1,False)
Call SetTextbox("To-currency","SVALD-VALUE","",DT_OB08_0300_TOCURRENCY_OCC1,False)
Call SetTextbox("Valid from","SVALD-VALUE","",DT_OB08_0300_VALID_FROM_OCC1,False)

Call ClickButton("Continue   \(Enter\)",True)  
'Call TakeScreenShot

Call GetTableCellData("SAPL0SAPTCTRL_V_TCURR", "Dir.quot.", 1, "", "", "DT_OB08_OUTPUT", False)
Call ClickButton("Back   \(F3\)",False)
Wait 5
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

''----------------------Tcode FB01----------------------------

Call SetTcode(DT_OB08_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_OB08_0100_POSTING_DATE),False)
Call SetTextbox("Type","BKPF-BLART","",DT_OB08_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_OB08_0100_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_OB08_0100_DOCUMENT_DATE),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_OB08_0100_PERIOD,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_OB08_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_OB08_0100_REFERENCE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_OB08_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_OB08_0100_ACCOUNT,False)

Call TakeScreenShot
Call PressEnter
Call TakeScreenShot
Call PressEnter

Call SetTextbox("Amount","BSEG-WRBTR","",DT_OB08_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_TAX_CODE,False)
Call SetTextbox("Text","BSEG-SGTXT","","31102016151345",False)
Call ClickButton("All Acct Assignmts",False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_OB08_1012_BUSINESS_AREA,True)
Call SetTextbox("Transactn type","COBL-RMVCT","",DT_TRANSACTION_TYPE,False)
Call TakeScreenShot
Call PressEnter
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_OB08_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","1",DT_OB08_0300_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter
Call TakeScreenShot
Call PressEnter

Call SetTextbox("Amount","BSEG-WRBTR","",DT_OB08_0300_AMOUNT_OCC1,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_OB08_0300_TAX_CODE,False)
Call TakeScreenShot
Call PressEnter
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_OB08_1007_COST_CENTER,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_OB08_1012_BUSINESS_AREA,False)

Call TakeScreenShot
Call PressEnter

Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_NO_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NO_OUTPUT&" was posted in company code RO02")
Call TakeScreenShot

'Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_OB08_0100_DOCUMENT_DATE_OCC1),False)
Call SetTextbox("Document Date","BKPF-BLDAT","",DT_OB08_0100_DOCUMENT_DATE_OCC1,False)
''Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_OB08_0100_POSTING_DATE_OCC1),False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",DT_OB08_0100_POSTING_DATE_OCC1,False)
Call SetTextbox("Type","BKPF-BLART","",DT_OB08_0100_TYPE_OCC1,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_OB08_0100_COMPANY_CODE_OCC1,False)

Call SetTextbox("Period","BKPF-MONAT","",DT_OB08_0100_PERIOD_OCC1,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_OB08_0100_CURRENCYRATE_OCC1,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_OB08_0100_REFERENCE_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_OB08_0100_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_OB08_0100_ACCOUNT_OCC1,False)

Call TakeScreenShot
Call PressEnter
Call TakeScreenShot
Call PressEnter

Call SetTextbox("Amount","BSEG-WRBTR","",DT_OB08_0300_AMOUNT_OCC2,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_OB08_0300_TAX_CODE,False)
Call SetTextbox("Text","BSEG-SGTXT","","31102016151345",False)
Call ClickButton("All Acct Assignmts",False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_OB08_1012_BUSINESS_AREA_OCC1,True)
Call SetTextbox("Transactn type","COBL-RMVCT","",DT_TRANSACTION_TYPE,False)
Call PressEnter()
Call TakeScreenShot
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_OB08_0300_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","1",DT_OB08_0300_ACCOUNT_OCC1,False)
Call TakeScreenShot
Call PressEnter
Call TakeScreenShot
Call PressEnter
Call SetTextbox("Amount","BSEG-WRBTR","",DT_OB08_0300_AMOUNT_OCC3,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_OB08_0300_TAX_CODE,False)
Call TakeScreenShot
Call PressEnter
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_OB08_1007_COST_CENTER_OCC1,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_OB08_1007_BUSINESS_AREA_OCC1,False)
Call TakeScreenShot
Call PressEnter

Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_NO_21_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NO_21_OUTPUT&" was posted in company code RO02")
Call TakeScreenShot

Call ClickButton("Back   \(F3\)",False)
Wait 5
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

''----------------------Tcode FBL3N----------------------------

Call SetTcode(DT_OB08_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_OB08_1000_GL_ACCOUNT,False)
Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)

Call ActivateNodeGuiTree(0, "Document;Document Number")
Wait 5
Call ClickButton("%_%%DYN007_%_APP_%-VALU_PUSH",False)
Wait 2
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_DOC_NO_OUTPUT, true)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_DOC_NO_21_OUTPUT, true)
Call ClickButton("Copy   \(F8\)",True)
Wait 2
Call ClickButton("Execute   \(F8\)",False)

''Call VerifyifGuiLabelExists(DT_AMOUNT_IN_RON_1)
''Call VerifyifGuiLabelExists(DT_AMOUNT_IN_RON_2)

Call ClickButton("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call TakeScreenShot

Call ClickButton("Find",True)

Call SetTextbox("Find","GD_SEARCHSTR","",DT_OB08_0850_FIND,True)

Call ClickButton("Continue   \(Enter\)",True)  
Wait 2
Call ClickButton("Show Sel\. Fields \(Ctrl\+F3\)",True)
Call ClickButton("Copy   \(Enter\)",True)
Call SetHorizontalScrollBar(30,False)
'Call VerifyifGuiLabelExists(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME)
Call VerifyifGuiLabelExists(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC1)

Call ClickButton("Back   \(F3\)",False)
Wait 5
Call ClickButton("Back   \(F3\)",False)
Wait 5
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)
''
''----------------------Tcode F-03----------------------------

Call SetTcode(DT_OB08_0100_OKCD_OCC2)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_OB08_0100_OKCD_OCC2)

Call SelectradioBUtton("RF05A-XPOS1","Document Number",False)

Call SetTextbox("Account","RF05A-AGKON","",DT_OB08_0131_ACCOUNT,False)
Call SetTextbox("Clearing Date","BKPF-BUDAT","",DT_OB08_0131_CLEARING_DATE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_OB08_0131_COMPANY_CODE,False)
Call SetTextbox("Period","BKPF-MONAT","",DT_OB08_0131_PERIOD,False)
Call SetTextbox("Currency","BKPF-WAERS","",DT_OB08_0131_CURRENCY,False)

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Call GetTextStatusBar("DT_STATUSBAR_TEXT")
If instr(1,DT_STATUSBAR_TEXT,"Period",1)  = 1 Then
	Call ClickButton("Process Open Items   \(Shift\+F4\)",False)	
End If
'if SAPGuiSession("Session").SAPGuiWindow("Clear G/L Account: Header").SAPGuiEdit("Period").exist(5) Then
'Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
'End If

Call SetTextbox("From","RF05A-SEL01",0,DT_DOC_NO_OUTPUT,False)
Call SetTextbox("From","RF05A-SEL01",1,DT_DOC_NO_21_OUTPUT,False)
''Call SetTextbox("From","RF05A-SEL01",0,DT_DOC_NO,False)
''Call SetTextbox("From","RF05A-SEL01",1,DT_DOC_NO_21,False)
Call ClickButton("Process Open Items   \(Shift\+F4\)",False)

Call SelectMenuBar("Document;Simulate")
''''
Call ClickButton("Choose   \(F2\)",False)
Call SetTextbox("Requested line item","\*BSEG-BUZEI",0,DT_LINE_ITEM,False)
Call ClickButton("Continue   \(Enter\)",False)
Call TakeScreenShot
Call SetTextbox("Text","BSEG-SGTXT","",DT_TEXT,False)
Call ClickButtonIfExist("All acct assignments",False)
'Call ClickButton("All Acct Assignmts",False)

Call SetTextbox("Transactn type","COBL-RMVCT","",DT_TRANSACTION_TYPE,False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Next Item   \(Shift\+F7\)",False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_TEXT,False)
Call ClickButtonIfExist("All acct assignments",False) ' Existing R1E code
Call ClickButtonIfExist("All Acct Assignmts",False)	 'Updating as Property got changed during upgrade

Call SetTextbox("Transactn type","COBL-RMVCT","",DT_TRANSACTION_TYPE,False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot
''''

Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_NO_3_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NO_3_OUTPUT&" was posted in company code RO02")
Call TakeScreenShot

Call ClickButton("Back   \(F3\)",False)
Wait 5
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

''----------------------Tcode F-03----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_OB08_0100_OKCD_OCC3)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC8)

Call SetTextbox("Document Number","RF05L-BELNR","",DT_OB08_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_OB08_0100_COMPANY_CODE_OCC2,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_OB08_0100_FISCAL_YEAR,False)

Call PressEnter()     
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("", 1, "Account", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

'Call VerifyGridCellContent("", 1, "Amount", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
'Call VerifyGridCellContent("", 2, "Amount", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)

Call VerifyGridCellContent("", 1, "Currency", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)

Call ClickButton("Back   \(F3\)",False)
Wait 5
Call ClickButton("Back   \(F3\)",False)
Wait 5
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

''----------------------Tcode FBL3N----------------------------

Call SetTcode(DT_OB08_0100_OKCD_OCC4)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC10)

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_OB08_1000_GL_ACCOUNT_OCC1,False)

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Wait 5
Call ActivateNodeGuiTree("","Document;Document Number")
Wait 5
Call ClickButton("%_%%DYN007_%_APP_%-VALU_PUSH",False)

Call SetTableDataNoRef("SAPLALDBSINGLE", "Single value", 1, DT_DOC_NO_OUTPUT, True)
Call SetTableDataNoRef("SAPLALDBSINGLE", "Single value", 2, DT_DOC_NO_21_OUTPUT, True)
Call SetTableDataNoRef("SAPLALDBSINGLE", "Single value", 3, DT_DOC_NO_3_OUTPUT, True)
Call TakeScreenShot
Call ClickButton("Copy   \(F8\)",True)
Call TakeScreenShot

Call SelectRadioBUtton("X_AISEL","All items",False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
''Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
''Call SelectRowGuiTable("SAPLSKBHTC_FIELD_LIST_820","Content","Document Number",True)
''Call ClickButton("Show sel\. fields \(CTRL\+F3\)",True)
''Call ClickButton("Copy   \(Enter\)",True)
''Call ClickButton("%_%%DYN001_%_APP_%-VALU_PUSH",True)
''Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_DOC_NO,True)
''Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_DOC_NO_2,True)
''Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",3,DT_DOC_NO_21,True)
''Call ClickButton("Copy   \(F8\)",True)
''Call ClickButton("Execute   \(Enter\)",True)
''Call TakeScreenShot
Call LogOff()
Call FinalStatus()


