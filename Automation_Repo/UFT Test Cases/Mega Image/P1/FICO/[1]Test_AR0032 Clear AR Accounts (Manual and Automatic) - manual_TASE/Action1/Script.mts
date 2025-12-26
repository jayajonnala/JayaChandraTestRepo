
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AR0032 Clear AR Accounts (Manual and Automatic) - manual_TASE
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


gstrTestCaseName = "Test_AR0032 Clear AR Accounts (Manual and Automatic) - manual_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AR0001 Manage AR Documents-Manage Manual Customer Invoicing_Cred.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''--------------login----------------'''''

''''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


''''----------------------Tcode OB08----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

'''No_Of_Record = Replace(SAPGuiSession("transaction:=OB08").SAPGuiWindow("transaction:=OB08").SAPGuiEdit("name:=VIM_POSITION_INFO").GetROProperty("value"),"Entry 1","")

Call ClickButton("VIM_POSI_PUSH",False)

Call SetTextboxNoLabel("SVALD-VALUE","",DT_OB08_0300_EXCH_RATE_TYPE,True)
Call SetTextbox("From currency","SVALD-VALUE","",DT_OB08_0300_FROM_CURRENCY,True)
Call SetTextbox("To-currency","SVALD-VALUE","",DT_OB08_0300_TOCURRENCY,True)
Call SetTextbox("Valid from","SVALD-VALUE","",ConvertDate(DT_OB08_0300_VALID_FROM),True)

'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

'''Posi_Row = Replace(Replace(Replace(SAPGuiSession("transaction:=OB08").SAPGuiWindow("transaction:=OB08").SAPGuiEdit("name:=VIM_POSITION_INFO").GetROProperty("value"),No_Of_Record,""),"Entry ",""),".","")
Call GetTextBoxvalue("VIM_POSITION_INFO",0,"DT_ROW_OUTPUT_2",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
''''Call GetTableCellData("SAPL0SAPTCTRL_V_TCURR","Dir.quot.",Posi_Row,"ValidFrom",DT_OB08_0300_VALID_FROM,"DT_OB08_0020_CHECK_TEXT_OF_TABLECELL_DIRQUOT_0_OUTPUT",False)
Call GetTableCellData("SAPL0SAPTCTRL_V_TCURR","Dir.quot.",DT_ROW_EXRT,"From","EUR","DT_OB08_0020_CHECK_TEXT_OF_TABLECELL_DIRQUOT_0_OUTPUT",False)


Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


Call ClickButton("VIM_POSI_PUSH",False)

Call SetTextboxNoLabel("SVALD-VALUE","",DT_OB08_0300_EXCH_RATE_TYPE_OCC1,True)
Call SetTextbox("From currency","SVALD-VALUE","",DT_OB08_0300_FROM_CURRENCY_OCC1,True)
Call SetTextbox("To-currency","SVALD-VALUE","",DT_OB08_0300_TOCURRENCY_OCC1,True)
Call SetTextbox("Valid from","SVALD-VALUE","",DT_OB08_0300_VALID_FROM_OCC1,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()
Call GetTextBoxvalue("VIM_POSITION_INFO",0,"DT_ROW_OUTPUT_3",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''''Posi_Row = Replace(Replace(Replace(SAPGuiSession("transaction:=OB08").SAPGuiWindow("transaction:=OB08").SAPGuiEdit("name:=VIM_POSITION_INFO").GetROProperty("value"),No_Of_Record,""),"Entry ",""),".","")

'''Call GetTableCellData("SAPL0SAPTCTRL_V_TCURR","Dir.quot.",Posi_Row,"ValidFrom",DT_OB08_0300_VALID_FROM_OCC1,"DT_OB08_0020_CHECK_TEXT_OF_TABLECELL_DIRQUOT_0_OCC1_OUTPUT",False)
Call GetTableCellData("SAPL0SAPTCTRL_V_TCURR","Dir.quot.",DT_ROW_EXRT_1,"From","EUR","DT_OB08_0020_CHECK_TEXT_OF_TABLECELL_DIRQUOT_0_OCC1_OUTPUT",False)


Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)


''----------------------Tcode F-22----------------------------
'Enter the Tcode
Call SetTcode(DT_OB08_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",Replace((DT_OB08_0100_DOCUMENT_DATE),"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_OB08_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_OB08_0100_COMPANY_CODE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_OB08_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_OB08_0100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_OB08_0100_DOCHEADER_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_OB08_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_OB08_0100_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()  

Call SelectCheckbox("BKPF-XMWST","1",DT_OB08_0301_CALCULATE_TAX,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_OB08_0301_AMOUNT,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_OB08_0301_ASSIGNMENT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_OB08_0301_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_OB08_0301_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_OB08_0301_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()  
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

Call SetTextbox("Amount","BSEG-WRBTR","",DT_OB08_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_OB08_0300_TAX_CODE,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_OB08_0300_ASSIGNMENT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_OB08_0300_TEXT,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_OB08_1006_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_OB08_1006_COST_CENTER,False)
Call SetTextbox("Profit Center","COBL-PRCTR","",DT_OB08_1006_PROFIT_CENTER,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

'veryfy sattus bar content
Call GetStatusBar("item1","DT_DOC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_OB08_0100_CHECK_TEXT_OF_STATUSBAR)

Wait(1)
Call SetTextbox("Document Date","BKPF-BLDAT","",Replace((DT_OB08_0100_DOCUMENT_DATE_OCC1),"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_OB08_0100_TYPE_OCC1,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_OB08_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_OB08_0100_CURRENCYRATE_OCC1,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_OB08_0100_REFERENCE_OCC1,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_OB08_0100_DOCHEADER_TEXT_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_OB08_0100_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_OB08_0100_ACCOUNT_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()  

Call SelectCheckbox("BKPF-XMWST","1",DT_OB08_0301_CALCULATE_TAX_OCC1,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_OB08_0301_AMOUNT_OCC1,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_OB08_0301_ASSIGNMENT_OCC1,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_OB08_0301_TEXT_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_OB08_0301_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_OB08_0301_ACCOUNT_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()  
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

Call SetTextbox("Amount","BSEG-WRBTR","",DT_OB08_0300_AMOUNT_OCC1,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_OB08_0300_TAX_CODE_OCC1,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_OB08_0300_ASSIGNMENT_OCC1,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_OB08_0300_TEXT_OCC1,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_OB08_1006_BUSINESS_AREA_OCC1,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_OB08_1006_COST_CENTER_OCC1,False)
Call SetTextbox("Profit Center","COBL-PRCTR","",DT_OB08_1006_PROFIT_CENTER_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

'veryfy sattus bar content
Call GetStatusBar("item1","DT_DOC2_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_OB08_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)

Wait(1)
Call ClickButton("Back   \(F3\)",False)
Call ClickButtonIfExist("Yes",True)
'Capture the screenshot
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

'''----------------------Tcode FB03----------------------------
'
'Enter the Tcode
Call SetTcode(DT_OB08_0100_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_OB08_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FISCAL_YEAR,False)

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call GetTextboxValue("BKPF-XBLNR","","DT_OB08_0750_CHECK_TEXT_OF_REFERENCE_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)

'
'''----------------------Tcode F-32----------------------------
'
'Enter the Tcode
Call SetTcode(DT_OB08_0100_OKCD_OCC2) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRadioButton("RF05A-XPOS1",DT_OB08_0131_DOCUMENT_NUMBER,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_OB08_0131_ACCOUNT,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_OB08_0131_COMPANY_CODE,False)
Call SetTextbox("Currency","BKPF-WAERS","",DT_OB08_0131_CURRENCY,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("From","RF05A-SEL01","0",DT_OB08_0731_FROM,False)
Call SetTextbox("From","RF05A-SEL01","1",DT_OB08_0731_FROM_OCC1,False)

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS","Res.Items",False)
Call SetTableDataNoRef("SAPDF05XTC_6106","Residual Items",1,DT_OB08_6106_TABLECELL_RESIDUAL_ITEMS_0,False)
Call PressEnter() 
Call PressEnter() 
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Wait(2)
Call SelectMenuBar("Goto;First item")
Wait(1)
Call SetTextbox("Text","BSEG-SGTXT","",DT_OB08_0301_TEXT_OCC2,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
'Validate If invoice is generated
Call GetStatusBar("item1","DT_DOC3_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_OB08_0131_CHECK_TEXT_OF_STATUSBAR)

Wait(1)
Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC7)

''----------------------Tcode FB03----------------------------

'Enter the Tcode
Call SetTcode(DT_OB08_0100_OKCD_OCC3) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC8)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_OB08_0100_DOCUMENT_NUMBER_OCC1,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_OB08_0100_COMPANY_CODE_OCC2,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_OB08_0100_FISCAL_YEAR,False)

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Reference","BKPF-XBLNR","",DT_OB08_0750_CHECK_TEXT_OF_REFERENCE_OCC1,False)

Call VerifyGridCellContent("",1,"Posting Key","",DT_CHECK_PSTKY_1)
Call VerifyGridCellContent("",2,"Posting Key","",DT_CHECK_PSTKY_2)
'Call VerifyGridCellContent("",3,"Posting Key","",DT_CHECK_PSTKY_3)

Call VerifyGridCellContent("",1,"Assignment","",DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call VerifyGridCellContent("",2,"Assignment","",DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)

'Call VerifyGridCellContent("",1,"DMBTR","",DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR)
'Call VerifyGridCellContent("",2,"DMBTR","",DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMBTR)
'Call VerifyGridCellContent("",3,"DMBTR","",DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMBTR)

Call VerifyGridCellContent("",1,"Account","",DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"Account","",DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
'Call VerifyGridCellContent("",3,"Account","",DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)

Call DoubleClickGuiGridCell("",0,1, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Acc","BSEG-HKONT",0,DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AUGBL, False)
Call VerifyTextBoxContent("Amount in LC","BSEG-DMBTR",0,DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR, False)
Call ClickButton("Back   \(F3\)", False)
'Call VerifyGridCellContent("",1,"G/L Account","",DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AUGBL)
Call DoubleClickGuiGridCell("",0,2, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Acc","BSEG-HKONT",0,DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AUGBL, False)
''Call VerifyTextBoxContent("Amount in LC","BSEG-DMBTR",0,DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMBTR, False)
Call ClickButton("Back   \(F3\)", False)
'Call VerifyGridCellContent("",2,"G/L Account","",DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AUGBL)
Call DoubleClickGuiGridCell("",0,3, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT",0,DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AUGBL, False)
Call ClickButton("Back   \(F3\)", False)
'Call VerifyGridCellContent("",3,"G/L Account","",DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AUGBL)

Call ClickButton("Display Document Header   \(F5\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call VerifyTextBoxContent("Document type","BKPF-BLART","",DT_OB08_1710_CHECK_TEXT_OF_DOCUMENT_TYPE,False)
Call ClickButton("Continue/Confirm   \(Enter\)",True)

Wait(1)
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC9)

'
'''----------------------Tcode FBL5N----------------------------
'Enter the Tcode
Call SetTcode(DT_OB08_0100_OKCD_OCC4) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC10)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_OB08_1000_CUSTOMER_ACCOUNT,False)

Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_OB08_0100_COMPANY_CODE_OCC2,False)

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
'
''Capture the screenshot
'Call TakeScreenShot()
'
''set filter criteria
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_OB08_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_OB08_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",3,DT_OB08_3010_TABLECELL_SINGLE_VALUE_2,True)

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
Call ClickButton("Change layout\.\.\.   \(Ctrl\+F8\)",False)
'Call SelectRowGuiGridbyRowNo("Column Set","",73,True)
Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST",10, True)
'Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Show Sel. Fields \(Ctrl\+F3\)",True)
Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST",7, True)
'Call SelectRowGuiGridbyRowNo("Column Set","",58,True)
Call ClickButton("Show Sel. Fields \(Ctrl\+F3\)",True)

Call ClickButton("Copy   \(Enter\)",True)
'Call ClickButton("Transfer   \(Enter\)",True)
Call TakeScreenShot()

'verify the details
Call VerifyifGuiLabelExists_ByIndex("S_LEDR",0)
Call VerifyifGuiLabelExists_ByIndex("S_LEDG",0)
Call VerifyifGuiLabelExists_ByIndex("S_LEDG",1)
Call VerifyifGuiLabelExists_ByIndex("S_LEDG",2)
'Call VerifyGridCellContent("",1,"Cleared/open items symbol","",DT_OB08_0120_CHECK_TOOLTIP_OF_NO_NAME)
'Call VerifyGridCellContent("",3,"Cleared/open items symbol","",DT_OB08_0120_CHECK_TOOLTIP_OF_NO_NAME_OCC1)
'Call VerifyGridCellContent("",4,"Cleared/open items symbol","",DT_OB08_0120_CHECK_TOOLTIP_OF_NO_NAME_OCC2)
'Call VerifyGridCellContent("",5,"Cleared/open items symbol","",DT_OB08_0120_CHECK_TOOLTIP_OF_NO_NAME_OCC3)

Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME,0)
Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC4,0)
Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC9,0)
Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC13,0)
'Call VerifyGridCellContent("",1,"Document Number","",DT_OB08_0120_CHECK_TEXT_OF_NO_NAME)
'Call VerifyGridCellContent("",3,"Document Number","",DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC4)
'Call VerifyGridCellContent("",4,"Document Number","",DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC9)
'Call VerifyGridCellContent("",5,"Document Number","",DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC13)

''''Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC1,0)
Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCCC1,0)
Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC5,0)
Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC10,0)
Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC14,0)
'Call VerifyGridCellContent("",1,"Amount in local currency","",DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC1)
'Call VerifyGridCellContent("",3,"Amount in local currency","",DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC5)
'Call VerifyGridCellContent("",4,"Amount in local currency","",DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC10)
'Call VerifyGridCellContent("",5,"Amount in local currency","",DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC14)

'Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_20A1,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_20A1_OCC1,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_20A1_OCC2,0)
'
''Call VerifyGridCellContent("",1,"Terms of Payment","",DT_OB08_0120_CHECK_TEXT_OF_20A1)
''Call VerifyGridCellContent("",3,"Terms of Payment","",DT_OB08_0120_CHECK_TEXT_OF_20A1_OCC1)
''Call VerifyGridCellContent("",5,"Terms of Payment","",DT_OB08_0120_CHECK_TEXT_OF_20A1_OCC2)
Call ClickButton("Last column   \(Ctrl\+F4\)",False)
Wait(15)

Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC3,0)
Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC7,0)
Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC12,0)
Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC17,0)
'Call VerifyGridCellContent("",1,"Amount in doc. curr.","",DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC3)
'Call VerifyGridCellContent("",3,"Amount in doc. curr.","",DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC7)
'Call VerifyGridCellContent("",4,"Amount in doc. curr.","",DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC12)
'Call VerifyGridCellContent("",5,"Amount in doc. curr.","",DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC17)
'
Wait(1)
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC11)

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




