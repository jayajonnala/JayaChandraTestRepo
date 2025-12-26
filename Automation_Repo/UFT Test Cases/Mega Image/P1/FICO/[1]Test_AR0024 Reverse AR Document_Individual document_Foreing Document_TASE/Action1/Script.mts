
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AR0024 Reverse AR Document_Individual document_Foreing Document
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


gstrTestCaseName = "Test_AR0024 Reverse AR Document_Individual document_Foreing Document"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'

'----------------------Tcode OB08----------------------------
''Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

'''No_Of_Record = Replace(SAPGuiSession("transaction:=OB08").SAPGuiWindow("transaction:=OB08").SAPGuiEdit("name:=VIM_POSITION_INFO").GetROProperty("value"),"Entry 1","")

Call ClickButton("VIM_POSI_PUSH",False)

Call SetTextboxNoLabel("SVALD-VALUE","",DT_OB08_0300_EXCH_RATE_TYPE,True)
Call SetTextbox("From currency","SVALD-VALUE","",DT_OB08_0300_FROM_CURRENCY,True)
Call SetTextbox("To-currency","SVALD-VALUE","",DT_OB08_0300_TOCURRENCY,True)
Call SetTextbox("Valid from","SVALD-VALUE","",DT_OB08_0300_VALID_FROM,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

''Posi_Row = Replace(Replace(Replace(SAPGuiSession("transaction:=OB08").SAPGuiWindow("transaction:=OB08").SAPGuiEdit("name:=VIM_POSITION_INFO").GetROProperty("value"),No_Of_Record,""),"Entry ",""),".","")
''iRow = Split(Posi_Row," ")
'Call GetTableCellData("SAPL0SAPTCTRL_V_TCURR","Dir.quot.",Posi_Row,"ValidFrom",DT_OB08_0300_VALID_FROM,"DT_OB08_0020_CHECK_TEXT_OF_TABLECELL_DIRQUOT_0_OUTPUT",False)
Call GetTextBoxvalue("VIM_POSITION_INFO",0,"DT_ROW_OUTPUT_2",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'''''Call GetTableCellData("SAPL0SAPTCTRL_V_TCURR","Dir.quot.",iRow(0),"","","DT_OB08_0020_CHECK_TEXT_OF_TABLECELL_DIRQUOT_0_OUTPUT",False)
Call GetTableCellData("SAPL0SAPTCTRL_V_TCURR","Dir.quot.",DT_ROW_EXRT,"","","DT_OB08_0020_CHECK_TEXT_OF_TABLECELL_DIRQUOT_0_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Output",DataRowSet)

'Call ClickButton("VIM_POSI_PUSH",False)
'
'Call SetTextbox("Exch\.rate type","SVALD-VALUE","",DT_OB08_0300_EXCH_RATE_TYPE_OCC1,True)
'Call SetTextbox("From currency","SVALD-VALUE","",DT_OB08_0300_FROM_CURRENCY_OCC1,True)
'Call SetTextbox("To-currency","SVALD-VALUE","",DT_OB08_0300_TOCURRENCY_OCC1,True)
'Call SetTextbox("Valid from","SVALD-VALUE","",DT_OB08_0300_VALID_FROM_OCC1,True)
''Capture the screenshot
'Call TakeScreenShot()
'Call ClickButton("Continue   \(Enter\)",True)
''Capture the screenshot
'Call TakeScreenShot()
'
'Posi_Row = Replace(Replace(Replace(SAPGuiSession("transaction:=OB08").SAPGuiWindow("transaction:=OB08").SAPGuiEdit("name:=VIM_POSITION_INFO").GetROProperty("value"),No_Of_Record,""),"Entry ",""),".","")
'
'Call GetTableCellData("SAPL0SAPTCTRL_V_TCURR","Dir.quot.",Posi_Row,"ValidFrom",DT_OB08_0300_VALID_FROM_OCC1,"DT_OB08_0020_CHECK_TEXT_OF_TABLECELL_DIRQUOT_0_OCC1_OUTPUT",False)
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
'''Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
'
'
'''----------------------Tcode FB70----------------------------
'Enter the Tcode
Call SetTcode(DT_OB08_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
'Capture the screenshot
Call TakeScreenShot()
'
'Call SetTextbox("Company Code","BKPF-BUKRS","",DT_OB08_1000_COMPANY_CODE,False)
'
'Call PressEnter()
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_OB08_1000_COMPANY_CODE)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call SetTextbox("Customer","INVFO-ACCNT","",DT_OB08_0510_CUSTOMER,False)
Call PressEnter()
Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_OB08_0510_INVOICE_DATE),False)
Call SetTextbox("Posting Date","INVFO-BUDAT","",ConvertDate(DT_OB08_0510_POSTING_DATE),False)
Call PressEnter()
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_OB08_0510_REFERENCE,False)
Call PressEnter()
Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_OB08_0550_HEADERTEXT,False)
Call TakeScreenShot
Call SelectTab("TS", "Basic data", blnIsItPopup)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_OB08_0510_AMOUNT,False)
Call SetTextbox("Amount","INVFO-WAERS","",DT_OB08_0510_AMOUNT_OCC1,False)
Call SetTextbox("Tax Amount","INVFO-WMWST","",DT_OB08_0510_TAX_AMOUNT,False)
'Call SetCombo("INVFO-MWSKZ","10 (Output tax 24 %)")
Call SetComboByKey("INVFO-MWSKZ", DT_OB08_0510_TAX_AMOUNT_OCC1)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 1, "", "", DT_OB08_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 1, "", "", DT_OB08_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
'Call SetTableData("SAPLFSKBTABLE", "Tax code", 1, "", "", DT_FB70_0100_TABLECELL_TAX_CODE_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", 1, "", "", DT_OB08_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", 1, "", "", DT_OB08_0100_TABLECELL_COST_CENTER_0, False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME,0)
Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC1,0)
Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC2,0)
Call ClickButton("Post   \(Ctrl\+S\)", False)
Call GetStatusBar("item1", "DT_DOC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar("Document "&DT_DOC1_OUTPUT&" was posted in company code RO02")
Call ClickButtonIfExist("Cancel   \(F12\)", True)

'''''''--------TransactionCode-/FB03----------''''

Call SetTcode(DT_OB08_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)

Call SetTextbox("Document Number","RF05L-BELNR","",DT_DOC_NO,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_OB08_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_OB08_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()

Call VerifyGridCellContent("", 1, "Account", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "Account", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 1, "Posting Key", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "Posting Key", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "Posting Key", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call DoubleClickGuiGridCell("",0, 1, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Acc","BSEG-HKONT",0,DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT, False)
Call VerifyTextBoxContent("Bline Date","BSEG-ZFBDT",0,ConvertDate(DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZFBDT), False)
Call ClickButton("Back   \(F3\)", False)
Call DoubleClickGuiGridCell("",0,2, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT",0,DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT, False)
Call ClickButton("Back   \(F3\)", False)
Call DoubleClickGuiGridCell("",0,3, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT",0,DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT, False)
Call ClickButton("Back   \(F3\)", False)
Call VerifyGridCellContent("", 1, "Tax Code", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("", 2, "Tax Code", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ)
Call VerifyGridCellContent("", 3, "Tax Code", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MWSKZ)
Call VerifyGridCellContent("", 1, "Description", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOBEZ)
Call VerifyGridCellContent("", 2, "Description", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOBEZ)
Call VerifyGridCellContent("", 3, "Description", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KOBEZ)
Call VerifyGridCellContent("", 1, "AZBET", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "AZBET", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 3, "AZBET", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)
Call VerifyGridCellContent("", 1, "Currency", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW)
Call VerifyGridCellContent("", 3, "Currency", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_RF05A_UBAZW)
Call ClickButton("Change Display Currency   \(F8\)", False)
Call SelectRadioButtonByIndexIfPopupExists("RF05L-XPOS1",1)
Call ClickButton("Continue/Confirm   \(Enter\)", True)
'''Call VerifyGridCellContent("", 1, "Amount", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR)
'''Call VerifyGridCellContent("", 2, "Amount", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMBTR)
'''Call VerifyGridCellContent("", 3, "Amount", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMBTR)
Call VerifyGridCellContent("", 1, "Amount", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 3, "Amount", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)

Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)
''
'''''''--------TransactionCode-/FB08----------''''

Call SetTcode(DT_OB08_0100_OKCD_OCC2)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)

Call SetTextbox("Document Number","RF05A-BELNS","",DT_DOC_NO,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_OB08_0105_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05A-GJAHS","",DT_OB08_0105_FISCAL_YEAR,False)
Call SetTextbox("Reversal Reason","UF05A-STGRD","",DT_OB08_0105_REVERSAL_REASON,False)
Call TakeScreenShot
Call ClickButton("Display document before reversal   \(F5\)", False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC7)
Call ClickButton("Back   \(F3\)", False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC8)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)", False)
Call GetStatusBar("item1", "DT_DOC2_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar("Document "&DT_DOC2_OUTPUT&" was posted in company code RO02")
Call ClickButtonIfExist("Cancel   \(F12\)", True)
'
'''''''--------TransactionCode-/FB03----------''''

Call SetTcode(DT_OB08_0100_OKCD_OCC3)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC10)

Call SetTextbox("Document Number","RF05L-BELNR","",DT_DOC_NO_2,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_OB08_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_OB08_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()

Call VerifyTextBoxContent("Posting Date","BKPF-BUDAT",0,ConvertDate(DT_PSTNG_DATE), False)
Call VerifyTextBoxContent("Document Date","BKPF-BLDAT",0,ConvertDate(DT_DOCU_DATE), False)
Call VerifyGridCellContent("", 1, "Account", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)
Call VerifyGridCellContent("", 2, "Account", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)
Call VerifyGridCellContent("", 3, "Account", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR_OCC1)
Call VerifyGridCellContent("", 1, "Posting Key", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
Call VerifyGridCellContent("", 2, "Posting Key", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC1)
Call VerifyGridCellContent("", 3, "Posting Key", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL_OCC1)
Call DoubleClickGuiGridCell("",0, 1, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("Amount in LC","BSEG-DMBTR",0,DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR_OCC1, False)
Call VerifyTextBoxContent("G/L Acc","BSEG-HKONT",0,DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT_OCC1, False)
Call VerifyTextBoxContent("Bline Date","BSEG-ZFBDT",0,ConvertDate(DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZFBDT_OCC1), False)
Call ClickButton("Display Additional Data   \(F8\)", False)
Call VerifyCheckBoxValue("BSEG-XNEGP","ON")
Call TakeScreenShot
Call ClickButton("Continue/Confirm   \(Enter\)", True)
Call ClickButton("Back   \(F3\)", False)
Call DoubleClickGuiGridCell("",0,2, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("Amt.in loc.cur.","BSEG-DMBTR",0,DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMBTR_OCC1, False)
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT",0,DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT_OCC1, False)
Call ClickButton("Display Additional Data   \(F8\)", False)
Call VerifyCheckBoxValue("BSEG-XNEGP","ON")
Call TakeScreenShot
Call ClickButton("Continue/Confirm   \(Enter\)", True)
Call ClickButton("Back   \(F3\)", False)
Call DoubleClickGuiGridCell("",0,3, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT",0,DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT_OCC1, False)
Call VerifyTextBoxContent("Amount in LC","BSEG-DMBTR",0,DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMBTR_OCC1, False)

Call ClickButton("Display Additional Data   \(F8\)", False)
Call VerifyCheckBoxValue("BSEG-XNEGP","ON")
Call TakeScreenShot
Call ClickButton("Continue/Confirm   \(Enter\)", True)
Call ClickButton("Back   \(F3\)", False)
Call VerifyGridCellContent("", 1, "Tax Code", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ_OCC1)
Call VerifyGridCellContent("", 2, "Tax Code", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ_OCC1)
Call VerifyGridCellContent("", 3, "Tax Code", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MWSKZ_OCC1)
Call VerifyGridCellContent("", 1, "Description", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOBEZ_OCC1)
Call VerifyGridCellContent("", 2, "Description", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOBEZ_OCC1)
Call VerifyGridCellContent("", 3, "Description", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KOBEZ_OCC1)
Call VerifyGridCellContent("", 1, "AZBET", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET_OCC1)
Call VerifyGridCellContent("", 2, "AZBET", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OCC1)
Call VerifyGridCellContent("", 3, "AZBET", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET_OCC1)
Call VerifyGridCellContent("", 1, "Currency", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW_OCC1)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW_OCC1)
Call VerifyGridCellContent("", 3, "Currency", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_RF05A_UBAZW_OCC1)
Call ClickButton("Change Display Currency   \(F8\)", False)
Call SelectRadioButtonByIndexIfPopupExists("RF05L-XPOS1",1)
Call ClickButton("Continue/Confirm   \(Enter\)", True)
'Call VerifyGridCellContent("", 1, "Amount", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR_OCC1)
'Call VerifyGridCellContent("", 2, "Amount", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMBTR_OCC1)
'Call VerifyGridCellContent("", 3, "Amount", 0, DT_OB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMBTR_OCC1)
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)

'''----------------------Tcode FBL5N----------------------------
'Enter the Tcode
Call SetTcode(DT_OB08_0100_OKCD_OCC4) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC12)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRadioButton("X_AISEL", "All items", False)
Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_OB08_1000_CUSTOMER_ACCOUNT,False)
Call ClickButton("Dynamic selections   \(Shift\+F4\)", False)
Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH", False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_OB08_3010_TABLECELL_SINGLE_VALUE_0, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_OB08_3010_TABLECELL_SINGLE_VALUE_1, True)
Call TakeScreenShot
Call ClickButton("Copy   \(F8\)", True)
Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot

'verify the details
Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC3,0)
Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC4,0)
'Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_ACCOUNT_11003416,0)
Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_NO_NAME_OCC5,0)
Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_RON,0)
Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_RON_OCC1,1)
'Call ClickButton("Change layout...   \(Ctrl\+F8\)", False)
'Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex(DT_DOC_DATE1,0)
Call VerifyifGuiLabelExists_ByIndex(DT_DOC_DATE1,1)
Call VerifyifGuiLabelExists_ByIndex("S_LEDG",0)
Call VerifyifGuiLabelExists_ByIndex("S_LEDG",1)
Call VerifyifGuiLabelExists_ByIndex(DT_DOC_TYPE1,0)
Call VerifyifGuiLabelExists_ByIndex(DT_DOC_TYPE2,1)
Call VerifyifGuiLabelExists_ByIndex(DT_CLRNG_DOC1,0)
Call VerifyifGuiLabelExists_ByIndex(DT_CLRNG_DOC2,1)
Call TakeScreenShot
Call ClickButton("Last column   \(Ctrl\+F4\)", False)
Call TakeScreenShot
'wait(10)
'
Call VerifyifGuiLabelExists_ByIndex(DT_OB08_0120_CHECK_TEXT_OF_ACCOUNT_11003416,0)
Call VerifyifGuiLabelExists_ByIndex(DT_EXP_EX_RATE,0)
Call VerifyifGuiLabelExists_ByIndex(DT_EXP_EX_RATE,1)
'Call TakeScreenShot

'Call ClickButton("Back   \(F3\)",False)
'Call ClickButton("Back   \(F3\)",False)


'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

