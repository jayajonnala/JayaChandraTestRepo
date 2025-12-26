		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.06.01.01.06 Manage Manual Customer Credit Note
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

gstrTestCaseName = "Test_09.06.01.01.06 Manage Manual Customer Credit Note"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''--------TransactionCode-FB75----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB75_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("Document date","INVFO-BLDAT", "", ConvertDate(DT_FB75_0510_DOCUMENT_DATE), False)
Call SetTextbox("Customer","INVFO-ACCNT", "", DT_FB75_0510_CUSTOMER, False)
Call SetTextbox("Text","INVFO-SGTXT", "", DT_FB75_0510_TEXT, False)
Call TakeScreenShot
Call SetComboByKey("Document type", DT_FB70_0510_DOCUMENT_TYPE)
Call TakeScreenShot
Call SetComboByKey("Tax Amount", DT_FB75_0510_TAX_AMOUNT)
Call TakeScreenShot

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB75_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB75_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_FB75_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_FB75_0100_TABLECELL_COST_CENTER_0, False)
Call PressEnter()
Call PressEnter()
Call TakeScreenShot

Call SelectTab("TS","Payment",False)
Call TakeScreenShot
Call SetTextbox("Bline Date","INVFO-ZFBDT", "", ConvertDate(DT_FB75_0520_BLINE_DATE), False)
Call TakeScreenShot
Call SelectTab("TS","Basic data",False)
Call TakeScreenShot
Call SelectCheckbox("INVFO-XMWST","0",DT_FB75_0510_CALCULATE_TAX,False)
Call TakeScreenShot

Call GetTextboxValue("RF05A-AZSAL",0,"DT_FB70_1200_CHECK_TEXT_OF_BAL_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_FB70_1200_CHECK_TEXT_OF_BAL_OUTPUT",DT_FB70_1200_CHECK_TEXT_OF_BAL)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Wait(5)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB75_0510_AMOUNT, False)
Call PressEnter()
Call TakeScreenShot

Call ClickCellTableByRowNo("SAPLFSKBTABLE", "Long Txt", 1, False)
Call ClickButton("Other languages",True)
Call SetTextbox("Language","SVALD-VALUE","",DT_FB75_0300_LANGUAGE,True)
Call PressEnter()

' GetspecialTextboxValue(textboxName, textboxIndex, dataTableColumnName, blnIsItPopup)
Call GetspecialTextboxValue("EENO_DYNP-BZCHNG",0,"DT_FB75_1001_CHECK_TEXT_OF_MEANING_OUTPUT",True)
Call TakeScreenShot
'Call WriteRunTimeDataToExcelGlobalSheet("DT_FB75_1001_CHECK_TEXT_OF_MEANING_OUTPUT",DT_FB75_1001_CHECK_TEXT_OF_MEANING)
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetSpecialTextbox("1st line","EENO_DYNP-ZEILE","",DT_FB75_1001_CHECK_TEXT_OF_MEANING_OUTPUT,True)
Call TakeScreenShot
Call ClickButton("Copy text   \(F5\)",True)
Call ClickButton("Simulate Document Posting   \(F9\)",False)
Call Pressenter()
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_FB75_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT")
Call VerifyStatusBar("Document "&DT_FB75_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB75_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT",DT_FB75_1000_CHECK_TEXT_OF_COMPANY_CODE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call ClickButtonIfExist("Cancel   \(F12\)",False)

''''''''--------TransactionCode-FBL5N ----------''''

Call SetTcode(DT_FB75_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FB75_1000_CUSTOMER_ACCOUNT,False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FB75_1106_DOCUMENT_NUMBER,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB75_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_FB75_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB75_0841_SEARCH_TERM_OCC1,True)
Call SetComboByKey("Search Direction",DT_FB75_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB75_0841_SEARCH_TERM_OCC2,True)
Call SetComboByKey("Search Direction",DT_FB75_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB75_0841_SEARCH_TERM_OCC3,True)
Call SetComboByKey("Search Direction",DT_FB75_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB75_0841_SEARCH_TERM_OCC4,True)
Call SetComboByKey("Search Direction",DT_FB75_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB75_0841_SEARCH_TERM_OCC5,True)
Call SetComboByKey("Search Direction",DT_FB75_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_FB75_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_FB75_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_FB75_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_FB75_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FB75_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_FB75_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("", 1, "ZTERM", 0, "")
Call VerifyGridCellContent("", 1, "KOSTL", 0, "")
Call VerifyGridCellContent("", 1, "BUDAT", 0, ConvertDate(DT_FB75_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT))
Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_FB75_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call TakeScreenShot

''''--------TransactionCode-F.62 ----------''''

Call SetTcode(DT_FB75_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectCheckbox("NORMBL", 0, DT_FB75_1000_STANDARD_DOCUMENTS, False)
Call SetTextbox("Fiscal year","RGJAHR-LOW","",Year(DT_FB75_1000_FISCAL_YEAR),False)
Call SetTextbox("Company code","RBUKRS-LOW","",DT_FB75_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Document number","RBELNR-LOW","",DT_FB75_1000_DOCUMENT_NUMBER,False)
Call SetTextbox("Document type","RBLART-LOW","",DT_FB75_1000_DOCUMENT_TYPE,False)
Call SetTextbox("Correspondence","REVENT","",DT_FB75_1000_CORRESPONDENCE,False)
Call SetTextbox("Posting date","RBUDAT-LOW","",ConvertDate(DT_FB75_1000_POSTING_DATE),False)
Call SetTextbox("Document date","RBLDAT-LOW","",ConvertDate(DT_FB75_1000_DOCUMENT_DATE),False)
Call SetTextbox("Delete if finished since","RDELDAYS","","",False)
'Call SetTextbox("Delete if finished since","RDELDAYS","",DT_FB75_1000_DELETE_IF_FINISHED_SINCE,False)
Call SetTextbox("Log to printer","PRDEST","",DT_FB75_1000_LOG_TO_PRINTER,False)
Call TakeScreenShot
Call ClickBUtton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Yes",True)
Call TakeScreenShot
'Call SetTextbox("Output to Printer","TDDEST","",DT_FB75_1000_LOG_TO_PRINTER,False)
Call PressEnter()

''''--------TransactionCode-SP02 ----------''''

Call SetTcode(DT_FB75_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

' SetFocusGuiLabel(labelContent, xCord, yCord, blnIsItPopup)
Call SetFocusGuiLabel(DT_FB75_0120_CHECK_TEXT_OF_NO_NAME,"347","56", False)
'Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenshot
Call VerifyifGuiLabelExistsByRelativeid(DT_FB75_0120_CHECK_TEXT_OF_NO_NAME,"wnd\[0\]/usr/lbl\[49,3\]")
Call TakeScreenshot
Call ClickLabel("X__PDF", "0", False)
Call TakeScreenshot

Call LogOff'
Call FinalStatus()
