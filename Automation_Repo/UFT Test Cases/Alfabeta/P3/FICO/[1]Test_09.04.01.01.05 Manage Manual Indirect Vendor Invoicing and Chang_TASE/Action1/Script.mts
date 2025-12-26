

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.01.01.05 Manage Manual Indirect Vendor Invoicing and Chang
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

gstrTestCaseName = "Test_09.04.01.01.05 Manage Manual Indirect Vendor Invoicing and Chang"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\jjonn\Desktop\TASEWork\Data\TASE_DT_09.04.01.01.02 Manage Manual Direct Foreign Vendor Invoicing.xls"
'strResultFolderPath = "C:\Users\jjonn\Desktop\TASEWork\Results"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''


'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''--------TransactionCode-FB60 ----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB60_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextboxNolabel("INVFO-ACCNT","",DT_FB60_0010_VENDOR,False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB60_0010_AMOUNT,False)
Call SetTextbox("Invoice date", "INVFO-BLDAT", "", ConvertDate(DT_FB60_0010_INVOICE_DATE), False)
Call SetTextbox("Posting Date", "INVFO-BUDAT", "", ConvertDate(DT_FB60_0010_POSTING_DATE), False)
Call SetTextbox("Reference", "INVFO-XBLNR", "", DT_FB60_0010_REFERENCE, False)
Call SetTextbox("Amount", "INVFO-WAERS", "",DT_FB60_0010_AMOUNT_OCC1 , False)
Call SetTextbox("Text", "INVFO-SGTXT", "", DT_FB60_0010_TEXT, False)
Call TakeScreenShot

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB60_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB60_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Text", "1", "", "", DT_FB60_0100_TABLECELL_TEXT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Tax Code", "1", "", "", DT_FB60_0100_TABLECELL_TAX_CODE_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_FB60_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_FB60_0100_TABLECELL_COST_CENTER_0, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call PressEnter()

Call SelectTab("TS","Payment",False)
Call TakeScreenShot

Call SetTextbox("Pmnt Terms", "INVFO-ZTERM", "", DT_FB60_0020_PAYT_TERMS, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call GetTextboxValue("INVFO-ZBD1T","","DT_FB60_0020_CHECK_TEXT_OF_DAYS_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB60_0020_CHECK_TEXT_OF_DAYS_OUTPUT",DT_FB60_0020_CHECK_TEXT_OF_DAYS)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetTextboxValue("INVFO-NETDT","","DT_FB60_0020_CHECK_TEXT_OF_DUE_ON_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB60_0020_CHECK_TEXT_OF_DUE_ON_OUTPUT",DT_FB60_0020_CHECK_TEXT_OF_DUE_ON)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call PressEnter()
Call TakeScreenShot

Call ClickButtonIfExist("Simulate Document Posting   \(F9\)",False)
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_1_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_1_OUTPUT&" was posted in company code GR02" )
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_1_OUTPUT",DT_DOC_1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call ClickButtonIfExist("Cancel   \(F12\)",False)

'''--------TransactionCode-FBL1N ----------''''

Call SetTcode(DT_FB60_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectRadioButton("X_AISEL","All items",False)
Call SetTextbox("Company Code","KD_BUKRS-LOW","",DT_FB60_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_FB60_1000_VENDOR_ACCOUNT,False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN012_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FB60_3010_TABLECELL_SINGLE_VALUE_0,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot

Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB60_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_FB60_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "FAEDT", 0, ConvertDate(DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_FAEDT))
Call LogOff'
Call FinalStatus()

