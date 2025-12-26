'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_PRE_04_E10G07P02S01V01 Rep on Gen Tax Docs from BELUX for BE10 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_PRE_04_E10G07P02S01V01 BELUX for BE10"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

'Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''--------TransactionCode-FB60----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB60_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call PressEnter()
Call SetComboByKey("RF05A-BUSCS",DT_FB60_1100_TRANSACTN)
Call SetTextboxNoLabel("INVFO-ACCNT","",DT_FB60_0010_VENDOR,False)
'This function SetTextboxNoLabel is used for vendor to supplier change.
Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_FB60_0010_INVOICE_DATE),False)
Call SetTextbox("Posting Date","INVFO-BUDAT","",ConvertDate(DT_FB60_0010_POSTING_DATE),False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
GetRowNo = 4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB60_0010_REFERENCE,False)   
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call SetComboByKey("Document type",DT_FB60_0010_DOCUMENT_TYPE_OCC1)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB60_0010_AMOUNT,False)
Call SetTextbox("Amount","INVFO-WAERS","",DT_FB60_0010_AMOUNT_OCC1,False)
Call SetTextbox("Tax Amount","INVFO-WMWST","",DT_FB60_0010_TAX_AMOUNT,False)
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 1, "", "", DT_FB60_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 1, "", "", DT_FB60_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Tax code", 1, "", "", DT_FB60_0100_TABLECELL_TAX_CODE_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", 1, "", "", DT_FB60_0100_TABLECELL_COST_CENTER_0, False)
Call TakeScreenShot
Call ClickButtonIfExist("Simulate Document Posting   \(F9\)", False)
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call VerifyifGuiLabelExists_ByIndex(DT_FB60_0120_CHECK_TEXT_OF_ND, 1)
Call VerifyifGuiLabelExists(DT_FB60_0120_CHECK_TEXT_OF_NO_NAME)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)", False)
Call GetStatusBar("item1", "DT_OP_DOC1")
GetRowNo = 4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call ClickButtonIfExist("Cancel   \(F12\)", True)



'''''''--------TransactionCode-/FB03----------''''


Call SetTcode(DT_FB60_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)


Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB60_0100_DOCUMENT_NUMBER,False)

Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB60_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB60_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "Tax Code", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("", 2, "Tax Code", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ)
Call VerifyGridCellContent("", 3, "Tax Code", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MWSKZ)
Call VerifyGridCellContent("", 2, "Profit Center", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOSTL)
Call VerifyGridCellContent("", 1, "Assignment", 0,DT_FB60_0750_CHECK_TEXT_OF_REFERENCE)
Call ClickButtonIfExist("Cancel   \(F12\)", True)


'''''--------TransactionCode-FB60----------''''

Call SetTcode(DT_FB60_0750_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB60_1000_COMPANY_CODE_OCC1,True)
Call TakeScreenShot
Call PressEnter()
Call SetComboByKey("RF05A-BUSCS",DT_FB60_1100_TRANSACTN_OCC1)
Call SetTextboxNoLabel("INVFO-ACCNT","",DT_FB60_0010_VENDOR_OCC1,False)
'This function SetTextboxNoLabel is used for vendor to supplier change.
Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_FB60_0010_INVOICE_DATE_OCC1),False)
Call SetTextbox("Posting Date","INVFO-BUDAT","",ConvertDate(DT_FB60_0010_POSTING_DATE_OCC1),False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
GetRowNo = 4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB60_0010_REFERENCE_OCC1,False)   
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call SetComboByKey("Document type",DT_FB60_0010_DOCUMENT_TYPE_OCC3)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB60_0010_AMOUNT_OCC2,False)
Call SetTextbox("Amount","INVFO-WAERS","",DT_FB60_0010_AMOUNT_OCC3,False)
Call SetTextbox("Tax Amount","INVFO-WMWST","",DT_FB60_0010_TAX_AMOUNT_OCC1,False)
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 1, "", "", DT_FB60_0100_TABLECELL_GL_ACCT_0_OCC1, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 1, "", "", DT_FB60_0100_TABLECELL_AMOUNT_IN_DOCCURR_0_OCC1, False)
Call SetTableData("SAPLFSKBTABLE", "Tax code", 1, "", "", DT_FB60_0100_TABLECELL_TAX_CODE_0_OCC1, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", 1, "", "", DT_FB60_0100_TABLECELL_COST_CENTER_0_OCC1, False)
Call TakeScreenShot
Call ClickButtonIfExist("Simulate Document Posting   \(F9\)", False)
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call VerifyifGuiLabelExists_ByIndex(DT_FB60_0120_CHECK_TEXT_OF_NE, 1)
Call VerifyifGuiLabelExists(DT_FB60_0120_CHECK_TEXT_OF_NO_NAME_OCC1)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)", False)
Call GetStatusBar("item1", "DT_OP_DOC2")
GetRowNo = 4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call ClickButtonIfExist("Cancel   \(F12\)", True)


'''''''--------TransactionCode-/FB03----------''''

Call SetTcode(DT_FB60_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)


Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB60_0100_DOCUMENT_NUMBER_OCC1,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB60_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB60_0100_FISCAL_YEAR_OCC1,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "Tax Code", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ_OCC1)
Call VerifyGridCellContent("", 2, "Tax Code", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ_OCC1)
Call VerifyGridCellContent("", 3, "Tax Code", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MWSKZ_OCC1)
Call VerifyGridCellContent("", 2, "Profit Center", 0, DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOSTL_OCC1)
Call VerifyGridCellContent("", 1, "Assignment", 0,DT_FB60_0750_CHECK_TEXT_OF_REFERENCE_OCC1)


'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
