
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_GL0031 Clear GL Accounts (Manual Automatic) - residual item manu_TASE
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


gstrTestCaseName = "Test_GL0031 Clear GL Accounts (Manual Automatic) - residual item manu_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'
''''''--------TransactionCode-FB60----------''''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB60_1000_COMPANY_CODE,True)
Call SetTextboxPopupIfExist("BKPF-BUKRS", "Company Code", DT_FB60_1000_COMPANY_CODE)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

'Call SetTextbox("Vendor","INVFO-ACCNT","",DT_FB60_0010_VENDOR,False)
Call SetTextboxNoLabel("INVFO-ACCNT","",DT_FB60_0010_VENDOR,False)
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Invoice date", "INVFO-BLDAT", "", ConvertDate(DT_FB60_0010_INVOICE_DATE), False)
Call SetTextbox("Posting Date", "INVFO-BUDAT", "", ConvertDate(DT_FB60_0010_POSTING_DATE), False)
Call PressEnter()     
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Reference", "INVFO-XBLNR", "", DT_FB60_0010_REFERENCE, False)
Call PressEnter()     
Call TakeScreenShot
Call SelectTab("TS","Payment",False)
Call TakeScreenShot
Call PressEnter()    
Call SetTextbox("Part\. bank","INVFO-BVTYP","",DT_FB60_0020_PART_BANK,False)
Call SetTextbox("Days","INVFO-ZBD1T","",DT_FB60_0020_DAYS,False)
Call SetTextbox("Pmnt Terms","INVFO-ZTERM","",DT_FB60_0020_PAYT_TERMS,False)
Call PressEnter()     
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call SelectTab("TS","Basic data",False)
Call TakeScreenShot
Call PressEnter()    
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB60_0010_AMOUNT,False)
Call SetTextbox("Amount","INVFO-WAERS","",DT_FB60_0010_AMOUNT_OCC1,False)
Call PressEnter()     
Call TakeScreenShot
Call PressEnter()    
Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB60_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", "*", False)
'newly added to handle banlancing profit center field error
Call SetTableData("SAPLFSKBTABLE", "Profit center", "1", "", "", DT_PROFIT_CENTER, False)
Call PressEnter
Call TakeScreenShot
Call PressEnter()    
Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call GetStatusBar("item1","DT_DOC_NO_1_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NO_1_OUTPUT&" was posted in company code RO02")
Call TakeScreenShot
Call ClickButtonIfExist("Cancel   \(F12\)",True)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NO_1_OUTPUT",DT_DOC_NO_1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''''''--------TransactionCode-F-02----------''''

Call SetTcode(DT_FB60_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_FB60_0100_POSTING_DATE),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB60_0100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB60_0100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB60_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB60_0100_ACCOUNT,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_FB60_0100_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB60_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FB60_0100_REFERENCE,False)
Call SetTextbox("Period","BKPF-MONAT","",MOnth(Date),False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
'Call PressEnter()
'Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB60_0300_AMOUNT,False)
Call SetTextbox("Tax code","BSEG-MWSKZ","",DT_FB60_0300_TAX_CODE,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_FB60_1007_COST_CENTER,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_FB60_1007_BUSINESS_AREA,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB60_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB60_0300_ACCOUNT,False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB60_0300_AMOUNT_OCC1,False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot


Call SelectMenuBar("Document;Simulate")
Call PressEnter()
Call TakeScreenShot

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_NO_2_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NO_2_OUTPUT&" was posted in company code RO02")

Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NO_2_OUTPUT",DT_DOC_NO_2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''
'''''''''''--------TransactionCode-F-03----------''''
'
Call SetTcode(DT_FB60_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)


Call SetTextbox("Account","RF05A-AGKON","",DT_FB60_0131_ACCOUNT,False)
Call SetTextbox("Clearing Date","BKPF-BUDAT","",ConvertDate(DT_FB60_0131_CLEARING_DATE),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_FB60_0131_PERIOD,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB60_0131_COMPANY_CODE,False)
Call SetTextbox("Currency","BKPF-WAERS","",DT_FB60_0131_CURRENCY,False)
Call TakeScreenShot
Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
While SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiEdit("guicomponenttype:=32","attachedtext:=Clearing Date","name:=BKPF-BUDAT","Index:=0").Exist(5)=True 
Call PressEnter()  
Wend

Call TakeScreenShot
Call ClickButton("Select All",False)
Call TakeScreenShot
''Call ClickButton("Activate Items",False)
Call ClickButton("Deactivate Items",False)
Call TakeScreenShot
Call ClickButton("Field content search",False)
Call TakeScreenShot

Call SelectRadioButton("RF05A-XPOS1", "Document Number", True)

Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("From","RF05A-SEL01",0,DT_DOC_NO_1,False)
Call SetTextbox("From","RF05A-SEL01",1,DT_DOC_NO_2,False)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call SelectTab("TS","Res.Items",False)

Call ClickButton("Select All",False)
Call TakeScreenShot
Call ClickButton("Activate Items",False)
Call TakeScreenShot
Call GetTextboxValue("RF05A-NETTO", "", " DT_FB60_6106_TABLECELL_RESIDUAL_ITEMS_OUTPUT", False)

Call SetTableDataNoRef("SAPDF05XTC_6106","Residual Items", 2, DT_FB60_6106_TABLECELL_RESIDUAL_ITEMS_OUTPUT, False)
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call Selectmenubar("Document;Simulate")

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call GetStatusBar("item1","DT_DOC_NO3_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NO3_OUTPUT&" was posted in company code RO02")

Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NO3_OUTPUT",DT_DOC_NO3)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''''''''''--------TransactionCode-FB03----------''''

Call SetTcode(DT_EXPECTEDTRANSACTIONCODE_OCC4)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB60_0100_DOCUMENT_NUMBER,False) 
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB60_0100_COMPANY_CODE_OCC1,False) 
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB60_0100_FISCAL_YEAR,False) 
Call PressEnter()     
Call TakeScreenShot


Call VerifyGridCellContent("", 1, "Account", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 2, "Account", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
'Call VerifyGridCellContent("", 3, "Account", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT)
Call VerifyGridCellContent("", 4, "Account", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_HKONT)

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
'Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "BSCHL", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 4, "BSCHL", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)

' Removing Lcase()
Call VerifyGridCellContent("", 1, "Description", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOBEZ)
Call VerifyGridCellContent("", 2, "Description", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOBEZ)
'Call VerifyGridCellContent("", 3, "Description", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KOBEZ)
Call VerifyGridCellContent("", 4, "Description", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KOBEZ)

Call VerifyGridCellContent("", 1, "Currency", 0, DT_FB60_0750_CHECK_TEXT_OF_CURRENCY)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW)
Call VerifyGridCellContent("", 3, "Currency", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_RF05A_UBAZW)

Call LogOff()
Call FinalStatus()
