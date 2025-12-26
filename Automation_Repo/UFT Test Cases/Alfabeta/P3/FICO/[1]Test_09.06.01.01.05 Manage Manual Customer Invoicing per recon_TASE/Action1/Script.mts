		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.06.01.01.05 Manage Manual Customer Invoicing per recon
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

gstrTestCaseName = "Test_09.06.01.01.05 Manage Manual Customer Invoicing per recon"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''''--------TransactionCode-FB70----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB70_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("Text","INVFO-SGTXT", "", DT_FB70_0510_TEXT, False)
Call SetTextbox("Invoice date","INVFO-BLDAT", "", ConvertDate(DT_FB70_0510_INVOICE_DATE), False)
Call SetTextbox("Customer","INVFO-ACCNT","",DT_FB70_0510_CUSTOMER,False)
Call SetComboByKey("INVFO-BLART",DT_FB70_0510_DOCUMENT_TYPE)
Call TakeScreenShot
Call SetComboByKey("INVFO-MWSKZ",DT_FB70_0510_TAX_AMOUNT)
Call TakeScreenShot
Call SelectCheckbox("INVFO-XMWST","0","ON",False)
Call TakeScreenShot
Call SetComboByKey("Transactn",DT_FB70_1200_TRANSACTN)
Call TakeScreenShot

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB70_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB70_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_FB70_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_FB70_0100_TABLECELL_COST_CENTER_0, False)
Call PressEnter()
Call TakeScreenShot
'
Call GetTextboxValue("RF05A-AZSAL",0,"DT_FB70_1200_CHECK_TEXT_OF_BAL_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_FB70_1200_CHECK_TEXT_OF_BAL_OUTPUT",DT_FB70_1200_CHECK_TEXT_OF_BAL)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Amount","INVFO-WRBTR", "", DT_FB70_0510_AMOUNT, False)
Call PressEnter()

Call SelectTab("TS","Details",False)
Call PressEnter()
Call TakeScreenShot
'Call VerifyTextBoxContent("G/L", "INVFO-HKONT", 0, DT_FB70_1200_CHECK_GL, False)

Call VerifyTextBoxEnabled("G/L", "INVFO-HKONT", 0,False)
Call TakeScreenShot
Call ClickButton("Simulate Document Posting   \(F9\)",False)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_FB70_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT")
Call VerifyStatusBar("Document "&DT_FB70_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet("DT_FB70_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT",DT_FB70_1000_CHECK_TEXT_OF_COMPANY_CODE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickButtonIfExist("Cancel   \(F12\)",False)

'''''--------TransactionCode-FBL5N ----------''''

Call SetTcode(DT_FB70_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FB70_1000_CUSTOMER_ACCOUNT,False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FB70_1000_CHECK_TEXT_OF_COMPANY_CODE,True)
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

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB70_0841_SEARCH_TERM_OCC5,True)
Call SetComboByKey("Search Direction",DT_FB70_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB70_0841_SEARCH_TERM_OCC6,True)
Call SetComboByKey("Search Direction",DT_FB70_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)

DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP = year(DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)& ConvertDoubledigit(CSTR(Month(DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)))&ConvertDoubledigit(CSTR(Day(DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)))&DT_FB70_0100_TABLECELL_COST_CENTER_0

Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP)
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("", 1, "BUDAT", 0, ConvertDate(DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT))
Call VerifyGridCellContent("", 1, "ZTERM", 0, DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZTERM)
Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_FB70_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call TakeScreenShot
Call LogOff()
Call FinalStatus ()



