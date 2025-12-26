		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.03.08 Post GL Adjustment (Ledger Specific)
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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
    RunTimeResultFolder= Parameter("RunTimeResultFolder")    
End If

gstrTestCaseName = "Test_09.07.01.03.08 Post GL Adjustment (Ledger Specific)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''''''--------TransactionCode-FB50L----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB50L_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("Document Date","ACGL_HEAD-BLDAT", "", ConvertDate(DT_FB50L_1010_DOCUMENT_DATE), False)
Call SetTextbox("Posting Date","ACGL_HEAD-BUDAT", "", ConvertDate(DT_FB50L_1010_POSTING_DATE), False)
Call SetTextbox("Currency","ACGL_HEAD-WAERS","",DT_FB50L_1010_CURRENCY,False)
Call SetTextbox("Ledger Grp","ACGL_HEAD-LDGRP","",DT_FB50L_1010_LEDGER_GRP,False)
Call SetTextbox("Reference","ACGL_HEAD-XBLNR","",DT_FB50L_1010_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","ACGL_HEAD-BKTXT","",DT_FB50L_1010_DOCHEADER_TEXT,False)
Call SetTextbox("Document type","ACGL_HEAD-BLART", "", DT_FB50L_1010_DOCUMENT_TYPE, False)


Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB50L_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "G/L acct", "2", "", "", DT_FB50L_0100_TABLECELL_GL_ACCT_1, False)

Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB50L_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "2", "", "", DT_FB50L_0100_TABLECELL_AMOUNT_IN_DOCCURR_1, False)

Call SelectCellGuiTable("SAPLFSKBTABLE", "D/C", "G/L acct", DT_FB50L_0100_TABLECELL_GL_ACCT_0, False)
Call SendKey("{F4}")
Wait 2
Call SendKey("{DOWN}")
Call SendKey("{TAB}")

Call SelectCellGuiTable("SAPLFSKBTABLE", "D/C", "G/L acct", DT_FB50L_0100_TABLECELL_GL_ACCT_1, False)
Call SendKey("{F4}")
Call SendKey("{TAB}")

Call SetTableData("SAPLFSKBTABLE", "Business area", "2", "", "", DT_FB50L_0100_TABLECELL_BUSINESS_AREA_1, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "2", "", "", DT_FB50L_0100_TABLECELL_COST_CENTER_1, False)
Call SetTableData("SAPLFSKBTABLE", "Profit center", "2", "", "", DT_FB50L_0100_TABLECELL_PROFIT_CENTER_1, False)
Call SetTableData("SAPLFSKBTABLE", "Trans.type", "1", "", "", DT_FB50L_0100_TABLECELL_TRANS_TYPE_0, False)

Call PressEnter
Call TakeScreenShot

Call ClickButtonIfExist("Simulate Document Posting   \(F9\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call GetStatusBar("item1","DT_DOC_1_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_1_OUTPUT&" was posted in company code GR02")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_1_OUTPUT",DT_DOC_1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call TakeScreenShot
Call ClickButtonIfExist("Cancel   \(F12\)",True)

'''''''''--------TransactionCode-FB50L----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB50L_1000_COMPANY_CODE_OCC1,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("Document Date","ACGL_HEAD-BLDAT", "", ConvertDate(DT_FB50L_1010_DOCUMENT_DATE_OCC1), False)
Call SetTextbox("Posting Date","ACGL_HEAD-BUDAT", "", ConvertDate(DT_FB50L_1010_POSTING_DATE_OCC1), False)
Call SetTextbox("Currency","ACGL_HEAD-WAERS","",DT_FB50L_1010_CURRENCY_OCC1,False)
Call SetTextbox("Ledger Grp","ACGL_HEAD-LDGRP","",DT_FB50L_1010_LEDGER_GRP_OCC1,False)
Call SetTextbox("Reference","ACGL_HEAD-XBLNR","",DT_FB50L_1010_REFERENCE_OCC1,False)
Call SetTextbox("Doc\.Header Text","ACGL_HEAD-BKTXT","",DT_FB50L_1010_DOCHEADER_TEXT_OCC1,False)
Call SetTextbox("Document type","ACGL_HEAD-BLART", "", DT_FB50L_1010_DOCUMENT_TYPE_OCC1, False)


Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB50L_0100_TABLECELL_GL_ACCT_0_OCC1, False)
Call SetTableData("SAPLFSKBTABLE", "G/L acct", "2", "", "", DT_FB50L_0100_TABLECELL_GL_ACCT_1_OCC1, False)

Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB50L_0100_TABLECELL_AMOUNT_IN_DOCCURR_0_OCC1, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "2", "", "", DT_FB50L_0100_TABLECELL_AMOUNT_IN_DOCCURR_1_OCC1, False)

Call SelectCellGuiTable("SAPLFSKBTABLE", "D/C", "G/L acct", DT_FB50L_0100_TABLECELL_GL_ACCT_0_OCC1, False)
Call SendKey("{F4}")
Wait 2
Call SendKey("{DOWN}")
Call SendKey("{TAB}")

Call SelectCellGuiTable("SAPLFSKBTABLE", "D/C", "G/L acct", DT_FB50L_0100_TABLECELL_GL_ACCT_1_OCC1, False)
Call SendKey("{F4}")
Call SendKey("{TAB}")

Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_FB50L_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_FB50L_0100_TABLECELL_COST_CENTER_0, False)
Call SetTableData("SAPLFSKBTABLE", "Profit center", "1", "", "", DT_FB50L_0100_TABLECELL_PROFIT_CENTER_0, False)

Call PressEnter
Call TakeScreenShot

Call ClickButtonIfExist("Simulate Document Posting   \(F9\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call GetStatusBar("item1","DT_DOC_2_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_2_OUTPUT&" was posted in company code GR02")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_2_OUTPUT",DT_DOC_2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call TakeScreenShot
Call ClickButtonIfExist("Cancel   \(F12\)",True)

'''''--------TransactionCode-FB03 ----------''''

Call SetTcode(DT_FB50L_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB50L_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB50L_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year", "RF05L-GJAHR", "", YEar(DT_FB50L_0100_FISCAL_YEAR), False)
Call PressEnter()     
Call TakeScreenShot

Call SelectMenuIdToolBar("&COL0",1)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB50L_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_FB50L_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyTextBoxContent("Ledger Group","BKPF-LDGRP", "", DT_FB50L_0750_CHECK_TEXT_OF_LEDGER_GROUP, False)

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("", 1, "KTONR", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call VerifyGridCellContent("", 1, "Amount", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)

Call VerifyGridCellContent("", 1, "Currency", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW)

Call VerifyGridCellContent("", 1, "Business Area", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GSBER)
Call VerifyGridCellContent("", 2, "Business Area", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_GSBER)

Call VerifyGridCellContent("", 1, "Profit Center", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOSTL)
Call VerifyGridCellContent("", 2, "Profit Center", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOSTL)

Call VerifyGridCellContent("", 1, "G/L Account", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 2, "G/L Account", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)


Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB50L_0100_DOCUMENT_NUMBER_OCC1,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB50L_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year", "RF05L-GJAHR", "", YEar(DT_FB50L_0100_FISCAL_YEAR_OCC1), False)
Call PressEnter()     
Call TakeScreenShot

Call SelectMenuIdToolBar("&COL0",1)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB50L_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_FB50L_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyTextBoxContent("Ledger Group","BKPF-LDGRP", "", DT_FB50L_0750_CHECK_TEXT_OF_LEDGER_GROUP_OCC1, False)

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC1)

Call VerifyGridCellContent("", 1, "KTONR", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)

Call VerifyGridCellContent("", 1, "Amount", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET_OCC1)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OCC1)

Call VerifyGridCellContent("", 1, "Currency", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW_OCC1)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW_OCC1)

Call VerifyGridCellContent("", 1, "Business Area", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GSBER_OCC1)
Call VerifyGridCellContent("", 2, "Business Area", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_GSBER_OCC1)

Call VerifyGridCellContent("", 1, "Profit Center", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOSTL_OCC1)
Call VerifyGridCellContent("", 2, "Profit Center", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOSTL_OCC1)

Call VerifyGridCellContent("", 1, "G/L Account", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT_OCC1)
Call VerifyGridCellContent("", 2, "G/L Account", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT_OCC1)
Call LogOff()
Call FinalStatus()

