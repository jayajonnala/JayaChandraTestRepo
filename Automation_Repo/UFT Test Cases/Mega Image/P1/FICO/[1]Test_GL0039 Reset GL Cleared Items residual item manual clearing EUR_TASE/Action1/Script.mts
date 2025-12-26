
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_GL0039 Reset GL Cleared Items residual item manual clearing EUR_TASE
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

gstrTestCaseName = "Test_GL0039 Reset GL Cleared Items residual item manual clearing EUR_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call StartDateof445PeriodByDate(DT_TODAY,"DT_STARTING_DATE_PERIOD")

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''''----------------------Tcode OB08----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call ClickButton("Position\.\.\.",False)
Call TakeScreenSHot()

'Call SetTextbox("Exch\.rate type","SVALD-VALUE","",DT_OB08_0300_EXCH_RATE_TYPE,False)
Call SetTextboxNoLabel("SVALD-VALUE","",DT_OB08_0300_EXCH_RATE_TYPE,False)

Call SetTextbox("From currency","SVALD-VALUE","",DT_OB08_0300_FROM_CURRENCY,False)
Call SetTextbox("To-currency","SVALD-VALUE","",DT_OB08_0300_TOCURRENCY,False)
Call SetTextbox("Valid from","SVALD-VALUE","",DT_OB08_0300_VALID_FROM,False)

Call ClickButton("Continue   \(Enter\)",True)     
'Call TakeScreenShot

Call GetTableCellData("SAPL0SAPTCTRL_V_TCURR", "Dir.quot.", 1, "", "", "DT_OB08_0020_CHECK_TEXT_OF_TABLECELL_DIRQUOT_0_OUTPUT", False)

Call ClickButton("Position\.\.\.",False)
Call TakeScreenSHot()

'Call SetTextbox("Exch\.rate type","SVALD-VALUE","",DT_OB08_0300_EXCH_RATE_TYPE_OCC1,False)
Call SetTextboxNoLabel("SVALD-VALUE","",DT_OB08_0300_EXCH_RATE_TYPE_OCC1,False)

Call SetTextbox("From currency","SVALD-VALUE","",DT_OB08_0300_FROM_CURRENCY_OCC1,False)
Call SetTextbox("To-currency","SVALD-VALUE","",DT_OB08_0300_TOCURRENCY_OCC1,False)
Call SetTextbox("Valid from","SVALD-VALUE","",DT_OB08_0300_VALID_FROM_OCC1,False)

Call ClickButton("Continue   \(Enter\)",True)  
'Call TakeScreenShot

Call GetTableCellData("SAPL0SAPTCTRL_V_TCURR", "Dir.quot.", 1, "", "", "DT_OB08_0020_CHECK_TEXT_OF_TABLECELL_DIRQUOT_0_OCC1_OUTPUT", False)
Call ClickBUtton("Back   \(F3\)",False)
Wait 5
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
'
''''''''''--------TransactionCode-F-02----------''''

Call SetTcode(DT_OB08_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

''Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_OB08_0100_POSTING_DATE),False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",DT_OB08_0100_POSTING_DATE,False)
''Call SetTextbox("Translation dte","BKPF-WWERT","",ConvertDate(DT_OB08_0100_TRANSLATN_DATE),False)
Call SetTextboxNolabel("BKPF-WWERT","",DT_OB08_0100_TRANSLATN_DATE,False) 'RGB
'Call SetTextbox("Translation dte","BKPF-WWERT","",DT_OB08_0100_TRANSLATN_DATE,False) 'R1E
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_OB08_0100_COMPANY_CODE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_OB08_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_OB08_0100_ACCOUNT,False)
''Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_OB08_0100_DOCUMENT_DATE),False)
Call SetTextbox("Document Date","BKPF-BLDAT","",DT_OB08_0100_DOCUMENT_DATE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_OB08_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_OB08_0100_REFERENCE,False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_OB08_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_OB08_0300_TAX_CODE,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_OB08_1007_COST_CENTER,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_OB08_1007_BUSINESS_AREA,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_OB08_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_OB08_0300_ACCOUNT,False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_OB08_0300_AMOUNT_OCC1,False)
Call ClickBUtton("All Acct Assignmts",False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_OB08_1012_BUSINESS_AREA,True)
Call TakeScreenShot
Call PressEnter

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call PressEnter()   
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC1_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC1_OUTPUT&" was posted in company code RO02")
Call TakeScreenShot


'''''''''--------TransactionCode-F-51----------''''

Call SetTcode(DT_OB08_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)

''Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_OB08_0122_DOCUMENT_DATE),False)
Call SetTextbox("Document Date","BKPF-BLDAT","",DT_OB08_0122_DOCUMENT_DATE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_OB08_0122_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_OB08_0122_COMPANY_CODE,False)
''Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_OB08_0122_POSTING_DATE),False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",DT_OB08_0122_POSTING_DATE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_OB08_0122_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_OB08_0122_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_OB08_0122_DOCHEADER_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_OB08_0122_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_OB08_0122_ACCOUNT,False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SetTextbox("Amount","BSEG-WRBTR","",DT_OB08_0300_AMOUNT_OCC2,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_OB08_0300_TAX_CODE_OCC1,False)
Call PressEnter()
Call TakeScreenShot
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_OB08_1007_COST_CENTER_OCC1,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_OB08_1007_BUSINESS_AREA_OCC1,False)

Call ClickBUtton("Choose open items   \(F6\)",False)
Call TakeScreenShot()

Call SetTextbox("Account Type","RF05A-AGKOA","",DT_OB08_0710_ACCOUNT_TYPE,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_OB08_0710_ACCOUNT,False)
Call SetTextbox("Company Code","RF05A-AGBUK","",DT_OB08_0710_COMPANY_CODE,False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButtoniFexist("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot

Call ClickButton("Select All",False)
Call TakeScreenShot
Call ClickButton("Activate Items",False)
Call TakeScreenShot
Call ClickBUtton("Field content search",False)
Call TakeScreenShot

Call SelectRadioBUtton("RF05A-XPOS1","Document Number",False)

Call ClickBUtton("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("From","RF05A-SEL01","",DT_DOC1_OUTPUT,False)

Call ClickBUtton("Continue   \(Enter\)",True)
Call TakeScreenShot

Call GetTextboxValue("RF05A-DIFFB","", "DT_NOT_ASSIGNED", False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SelectTab("TS","Res.Items",False)
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTableDataNoRef("SAPDF05XTC_6106","Residual Items",1,DT_OB08_6106_TABLECELL_RESIDUAL_ITEMS_0,False)

Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call PressEnter()   
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC2_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC2_OUTPUT&" was posted in company code RO02")
Call TakeScreenShot

''''''--------TransactionCode-F-51----------''''

Call SetTcode(DT_OB08_0100_OKCD_OCC2)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)

''Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_OB08_0122_DOCUMENT_DATE_OCC1),False)
Call SetTextbox("Document Date","BKPF-BLDAT","",DT_OB08_0122_DOCUMENT_DATE_OCC1,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_OB08_0122_COMPANY_CODE_OCC1,False)
''Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_OB08_0122_POSTING_DATE_OCC1),False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",DT_OB08_0122_POSTING_DATE_OCC1,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_OB08_0122_CURRENCYRATE_OCC1,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_OB08_0122_REFERENCE_OCC1,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_OB08_0122_DOCHEADER_TEXT_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_OB08_0122_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_OB08_0122_ACCOUNT_OCC1,False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_OB08_0300_AMOUNT_OCC3,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_OB08_0300_TAX_CODE_OCC2,False)
Call PressEnter()
Call TakeScreenShot
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_OB08_1007_COST_CENTER_OCC2,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_OB08_1007_BUSINESS_AREA_OCC2,False)

Call ClickBUtton("Choose open items   \(F6\)",False)
Call TakeScreenShot()

Call SetTextbox("Account Type","RF05A-AGKOA","",DT_OB08_0710_ACCOUNT_TYPE_OCC1,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_OB08_0710_ACCOUNT_OCC1,False)
Call SetTextbox("Company Code","RF05A-AGBUK","",DT_OB08_0710_COMPANY_CODE_OCC1,False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButtoniFexist("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot

Call ClickButton("Select All",False)
Call TakeScreenShot
Call ClickButton("Activate Items",False)
Call TakeScreenShot
Call ClickBUtton("Field content search",False)
Call TakeScreenShot

Call SelectRadioBUtton("RF05A-XPOS1","Document Number",False)

Call ClickBUtton("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("From","RF05A-SEL01","",DT_DOC2_OUTPUT,False)

Call ClickBUtton("Continue   \(Enter\)",True)
Call TakeScreenShot

Call GetTextboxValue("RF05A-DIFFB","", "DT_NOT_ASSIGNED_1", False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SelectTab("TS","Res.Items",False)
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTableDataNoRef("SAPDF05XTC_6106","Residual Items",1,DT_OB08_6106_TABLECELL_RESIDUAL_ITEMS_0_OCC1,False)

Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call PressEnter()
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC3_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC3_OUTPUT&" was posted in company code RO02")
Call TakeScreenShot

''''--------TransactionCode-FB03 ----------''''

Call SetTcode(DT_EXPECTEDTRANSACTIONCODE_OCC8)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_OB08_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_OB08_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year", "RF05L-GJAHR", "", DT_OB08_0100_FISCAL_YEAR, False)
Call PressEnter()     
Call TakeScreenShot


'''--------TransactionCode-FBRA ----------''''
Call SetTcode(DT_OB08_0100_OKCD_OCC4)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC10)

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Clearing Document","RF05R-AUGBL","",DT_OB08_0100_CLEARING_DOCUMENT,False)
Call SetTextbox("Company Code","RF05R-BUKRS","",DT_OB08_0100_COMPANY_CODE_OCC2,False)
Call SetTextbox("Fiscal Year","RF05R-GJAHR","",DT_OB08_0100_FISCAL_YEAR_OCC1,False)
Call TakeScreenShot
Call PressEnter

Call SelectMenuBar("Clearing;Reset cleared items")
'''Call ClickButtonIfExist("Resetting and revers", True)
Call ClickButtonIfExist("Reset and Reverse", True)
Call takeScreenShot()
Call ClickButtonIfExist("Yes",True)
Call SetTextbox("Reversal Reason","RF05R-STGRD","",DT_OB08_0300_REVERSAL_REASON,True)
Call SetTextbox("Posting Date","RF05R-BUDAT","",DT_OB08_0300_POSTING_DATE,True)
Call SetTextbox("Posting period","RF05R-MONAT","",DT_OB08_0300_POSTING_PERIOD,True)

Call Takescreenshot()

Call ClickButtonifExist("Continue   \(Enter\)",True)
Call Takescreenshot()
Call ClickButtonifExist("Continue   \(Enter\)",True)
Call Takescreenshot()
Call GetTextboxValue("MESSTXT1", 0, "DT_OB08_0020_CHECK_TEXT_OF_TABLECELL_DIRQUOT_0_OCC1_OUTPUT", True)

Call ClickButtonifExist("Continue   \(Enter\)",True)
Call Takescreenshot()


'''''--------TransactionCode-fbl3n----------''''

Call SetTcode(DT_OB08_0100_OKCD_OCC5)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC11)

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SelectRadioButton("X_AISEL","All items", False)

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_OB08_1000_GL_ACCOUNT,False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_OB08_1000_COMPANY_CODE,False)

Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ActivateNodeGuiTree(0, "Document;Document Number")
Call ClickButton("%_%%DYN007_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_OB08_3010_TABLECELL_SINGLE_VALUE_0,True)
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

'Call VerifyGridCellContent("", 83, "DMSHB", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_82_DMSHB)

Call SetFocusGuiLabel(DT_OB08_3010_TABLECELL_SINGLE_VALUE_0, "200", "120", False)
Call TakeScreenShot()

Call ClickButton("Display Document   \(Shift\+F2\)",False)
Call TakeScreenShot()

Call ClickButton("Call Up Document Overview   \(F9\)",False)
Call TakeScreenShot()

'''''--------TransactionCode-fbl3n----------''''

Call SetTcode(DT_OB08_0100_OKCD_OCC5)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC11)

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_OB08_1000_GL_ACCOUNT,False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_OB08_1000_COMPANY_CODE,False)

Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ActivateNodeGuiTree(0, "Document;Document Number")
Call ClickButton("%_%%DYN007_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_OB08_0100_CLEARING_DOCUMENT,True)
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call LogOff'
Call FInalStatus()


