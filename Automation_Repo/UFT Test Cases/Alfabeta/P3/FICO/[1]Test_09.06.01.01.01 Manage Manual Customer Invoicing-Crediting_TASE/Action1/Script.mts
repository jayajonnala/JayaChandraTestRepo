		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.06.01.01.01 Manage Manual Customer Invoicing-Crediting
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

gstrTestCaseName = "Test_09.06.01.01.01 Manage Manual Customer Invoicing-Crediting"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'

''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''''''--------TransactionCode-FB70----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB70_COMPANY_CODE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("Invoice date","INVFO-BLDAT", "", ConvertDate(DT_FB70_0510_INVOICE_DATE), False)
Call SetTextbox("Customer","INVFO-ACCNT","",DT_FB70_0510_CUSTOMER,False)
Call SetTextbox("Text","INVFO-SGTXT", "", DT_FB70_0510_TEXT, False)
Call TakeScreenShot
Call SetComboByKey("INVFO-MWSKZ",DT_FB70_0510_TAX_AMOUNT)
Call TakeScreenShot
Call SetComboByKey("INVFO-BLART",DT_FB70_0510_DOCUMENT_TYPE)
Call TakeScreenShot
Call SelectCheckbox("INVFO-XMWST","0",DT_FB70_0510_CALCULATE_TAX,False)
Call TakeScreenShot

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB70_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB70_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call PressEnter()
Call TakeScreenShot
'Call VerifyStatusBar(DT_FB70_1200_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot
Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_FB70_0100_TABLECELL_COST_CENTER_0, False)
Call TakeScreenShot
Call PressEnter()
Call VerifyTextBoxContent("Information Message","MESSTXT1", "0", Lcase(DT_FB70_0010_CHECK_TEXT_OF_MESSTXT1), True)
Call TakeScreenShot
Call ClickButtonifExist("Continue   \(Enter\)",True)
Call TakeScreenShot
Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_FB70_0100_TABLECELL_BUSINESS_AREA_0, False)
Call PressEnter()
Call TakeScreenShot
Call GetTextboxValue("RF05A-AZSAL","0", "DT_FB70_1200_CHECK_TEXT_OF_BAL_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_FB70_1200_CHECK_TEXT_OF_BAL_OUTPUT",DT_FB70_1200_CHECK_TEXT_OF_BAL)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Amount","INVFO-WRBTR", "", DT_FB70_0510_AMOUNT, False)
Call PressEnter()
Call TakeScreenShot

Call SelectTab("TS", "Details", False)
Call PressEnter()
Call TakeScreenShot

Call VerifyTextBoxEnabled("G/L","INVFO-HKONT",0,False)

Call ClickButton("Simulate Document Posting   \(F9\)",False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_FB70_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT")
Call VerifyStatusBar("Document "&DT_FB70_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet("DT_FB70_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT",DT_FB70_1000_CHECK_TEXT_OF_COMPANY_CODE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call ClickButtonIfExist("Cancel   \(F12\)",False)

''''''''--------TransactionCode-FBL5N----------''''

Call SetTcode(DT_FB70_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FB70_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_FB70_1000_COMPANY_CODE_OCC1,False)

Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FB70_1106_DOCUMENT_NUMBER,True)
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB70_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_FB70_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB70_0841_SEARCH_TERM_OCC1,True)
Call SetComboByKey("Search Direction",DT_FB70_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB70_0841_SEARCH_TERM_OCC2,True)
Call SetComboByKey("Search Direction",DT_FB70_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB70_0841_SEARCH_TERM_OCC3,True)
Call SetComboByKey("Search Direction",DT_FB70_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)


Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 1, "ZTERM", 0, DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZTERM)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("", 1, "BUDAT", 0, ConvertDate(DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT))

Call DoubleClickGuiGridCell("", "0", "1", "Document type", False)
Call TakeScreenShot
Call ClickButtonIfExist("Call Up Document Overview   \(F9\)",False)

Call ClickContextButtonToolBar("&MB_VARIANT", 1)
Call SelectMenuIdToolBar("&COL0",0)
Call ClickButtonToolBar("&FIND",0)

Call SetTextbox("Search Term:","GS_SEARCH-VALUE",0,"Cost Center",True)
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("",2,"Cost Center",0,DT_FB70_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOSTL)
Call TakeScreenShot


''''--------TransactionCode-F.62 ----------''''

Call SetTcode(DT_FB70_0750_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectCheckbox("NORMBL", 0, DT_FB70_1000_STANDARD_DOCUMENTS, False)
Call SetTextbox("Fiscal year","RGJAHR-LOW","",Year(DT_FB70_1000_FISCAL_YEAR),False)
Call SetTextbox("Company code","RBUKRS-LOW","",DT_FB70_1000_COMPANY_CODE_OCC2,False)
Call SetTextbox("Document number","RBELNR-LOW","",DT_FB70_1000_DOCUMENT_NUMBER,False)
Call SetTextbox("Document type","RBLART-LOW","",DT_FB70_1000_DOCUMENT_TYPE,False)
Call SetTextbox("Posting date","RBUDAT-LOW","",ConvertDate(DT_FB70_1000_POSTING_DATE),False)
Call SetTextbox("Correspondence","REVENT","",DT_FB70_1000_CORRESPONDENCE,False)
Call SetTextbox("Delete if finished since","RDELDAYS","","",False)
'Call SetTextbox("Delete if finished since","RDELDAYS","",DT_FB70_1000_DELETE_IF_FINISHED_SINCE,False)
Call SetTextbox("Log to printer","PRDEST","",DT_FB70_1000_LOG_TO_PRINTER,False)
Call ClickBUtton("Execute   \(F8\)",False)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Yes",True)
'Call SetTextbox("Output to Printer","TDDEST","",DT_FB70_1000_LOG_TO_PRINTER,False)
Call PressEnter()
wait (3)
Call TakeScreenShot

'''''--------TransactionCode-SP02 ----------''''

Call SetTcode(DT_SP02_TRN)     
Call PressEnter()     
Call TakeScreenShot

DT_DOC_TITLE_IP = DT_FB70_1000_DOCUMENT_NUMBER&"/"&Year(DT_FB70_1000_FISCAL_YEAR)
Call VerifyifGuiLabelExistsByRelativeid(DT_DOC_TITLE_IP,"wnd\[0\]/usr/lbl\[49,3\]")

Call SetFocusGuiLabel(DT_DOC_TITLE_IP,"347","56", False)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenshot
Call VerifyifGuiLabelExistsByRelativeid(DT_DOC_TITLE_IP,"wnd\[0\]/usr/lbl\[49,3\]")
Call TakeScreenshot
Call ClickLabel("X__PDF", "0", False)
Call TakeScreenshot

Call LogOff()
Call FinalStatus()



