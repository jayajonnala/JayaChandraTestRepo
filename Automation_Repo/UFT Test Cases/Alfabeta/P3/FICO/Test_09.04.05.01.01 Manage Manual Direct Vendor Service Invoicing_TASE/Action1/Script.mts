		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.05.01.01 Manage Manual Direct Vendor Service Invoicing
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

gstrTestCaseName = "Test_09.04.05.01.01 Manage Manual Direct Vendor Service Invoicing"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''''--------TransactionCode-FB65----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB65_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetComboByKey("Transactn",DT_FB65_1100_TRANSACTN)
Call TakeScreenShot

Call SetTextbox("Document date", "INVFO-BLDAT", "", ConvertDate(DT_FB65_0010_DOCUMENT_DATE), False)
Call SetTextboxNoLabel("INVFO-ACCNT","",DT_FB65_0010_VENDOR,False)
Call SetComboByKey("INVFO-BLART",DT_FB65_0010_DOCUMENT_TYPE)
Call TakeScreenShot
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB65_0010_AMOUNT,False)
Call SelectCheckbox("INVFO-XMWST","0","ON",False)
Call TakeScreenShot
Call SetComboByKey("INVFO-MWSKZ",DT_FB65_0010_INVFOMWSKZ)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Text", "INVFO-SGTXT", "", DT_FB65_0010_TEXT, False)
Call TakeScreenShot

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB65_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB65_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Text", "1", "", "", DT_FB65_0100_TABLECELL_TEXT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_FB65_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_FB65_0100_TABLECELL_COST_CENTER_0, False)
Call TakeScreenShot

Call ClickCellTableByRowNo("SAPLFSKBTABLE", "Long Txt", 1, False)
Call ClickButton("Other languages",True)
Call SetTextbox("Language","SVALD-VALUE","",DT_FB65_0300_LANGUAGE,True)
Call PressEnter()

Call SetTextbox("1st line","EENO_DYNP-ZEILE","",DT_FB65_1001_1ST_LINE,True)
Call ClickButton("Copy text   \(F5\)",True)
Call ClickButton("Simulate Document Posting   \(F9\)",False)
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_NUMBER_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NUMBER_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NUMBER_OUTPUT",DT_DOC_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call ClickButtonIfExist("Cancel   \(F12\)",False)


'''''--------TransactionCode-FBL1N ----------''''

Call SetTcode(DT_FB65_1100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectRadioButton("X_AISEL", "All items", FAlse)
Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_FB65_1000_VENDOR_ACCOUNT,False)
Call SetTextbox("Company code","KD_BUKRS-LOW","",DT_FB65_1000_COMPANY_CODE_OCC2,False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN012_%_APP_%-VALU_PUSH",False)
Call TakeScreenShot

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FB65_1106_DOCUMENT_NUMBER,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB65_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_FB65_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter

Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)


Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB65_0841_SEARCH_TERM_OCC1,True)
Call SetComboByKey("Search Direction",DT_FB65_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)


DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP = year(DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)& ConvertDoubledigit(CSTR(Month(DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)))& ConvertDoubledigit(CSTR(Day(DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)))
Call VerifyGridCellContent("", 1, "ICO_AUGP", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ICO_AUGP)
Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP)
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 1, "HWAER", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("", 1, "HWAER", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 1, "ZTERM", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZTERM)
Call VerifyGridCellContent("", 1, "HKONT", 0, DT_FB65_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call TakeScreenShot


''''--------TransactionCode-F.62 ----------''''

Call SetTcode(DT_FB65_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectCheckbox("NORMBL", 0, "ON", False)
Call SetTextbox("Fiscal year","RGJAHR-LOW","",Year(DT_FB65_1000_FISCAL_YEAR),False)
Call SetTextbox("Company code","RBUKRS-LOW","",DT_FB65_1000_COMPANY_CODE_OCC3,False)
Call SetTextbox("Document number","RBELNR-LOW","",DT_FB65_1000_DOCUMENT_NUMBER,False)
Call SetTextbox("Document type","RBLART-LOW","",DT_FB65_1000_DOCUMENT_TYPE,False)
Call SetTextbox("Correspondence","REVENT","",DT_FB65_1000_CORRESPONDENCE,False)
Call SetTextbox("Delete if finished since","RDELDAYS","","",False)
'Call SetTextbox("Delete if finished since","RDELDAYS","",DT_FB65_1000_DELETE_IF_FINISHED_SINCE,False)
Call SetTextbox("Log to printer","PRDEST","",DT_FB65_1000_LOG_TO_PRINTER,False)
Call ClickBUtton("Execute   \(F8\)",False)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Yes",True)

Call SetTextbox("Output Device","USR01-SPLD","",DT_FB65_1000_OUTPUT_TO_PRINTER,False)
Call PressEnter()
Call TakeScreenShot

'''''--------TransactionCode-SP02 ----------''''

Call SetTcode(DT_TRANSACTION)     
Call PressEnter()     
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
DT_DOC_TITLE_IP = DT_DOC_TITLE&"/"&Year(DT_FB65_1000_FISCAL_YEAR)
Call VerifyifGuiLabelExistsByRelativeid(DT_DOC_TITLE_IP,"wnd\[0\]/usr/lbl\[49,3\]")
Call ClickLabel("X__PDF", 0, False)
''added wiat time to load the PDF before taking screenshot
Wait 30
Call TakeScreenshot()

Call LogOff'
Call FinalStatus()


